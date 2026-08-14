using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Catalog;

public interface ICatalogSignatureVerifier
{
    bool Verify(ReadOnlySpan<byte> catalog, ReadOnlySpan<byte> signature);
}

public sealed class EcdsaCatalogSignatureVerifier : ICatalogSignatureVerifier, IDisposable
{
    private readonly ECDsa _key = ECDsa.Create();

    public EcdsaCatalogSignatureVerifier(string publicKeyPem)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(publicKeyPem);
        _key.ImportFromPem(publicKeyPem);
    }

    public bool Verify(ReadOnlySpan<byte> catalog, ReadOnlySpan<byte> signature) =>
        _key.VerifyData(catalog, signature, HashAlgorithmName.SHA256);

    public void Dispose() => _key.Dispose();
}

/// <summary>
/// Загружает только декларативные условия. Идентификаторы действий разрешаются локальным реестром,
/// поэтому каталог не может доставить исполняемый код.
/// </summary>
public sealed class KnownIssueCatalogService : IKnownIssueCatalog, IDisposable
{
    public const string PublicKeyPem = """
        -----BEGIN PUBLIC KEY-----
        MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE7UeU1usCRYHmk+j9SMgx0UsJ2PrW
        FzrSnMfxwoQG4U0a5Qc2DBhCrENP4yEAu8SSEOept2qicnR9rx8rRrDhcA==
        -----END PUBLIC KEY-----
        """;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };
    private static readonly HashSet<string> AllowedSourceHosts =
    [
        "learn.microsoft.com",
        "support.microsoft.com"
    ];

    private readonly HttpClient _httpClient;
    private readonly ICatalogSignatureVerifier _signatureVerifier;
    private readonly Uri _catalogUri;
    private readonly string _cacheDirectory;
    private KnownIssueCatalogSnapshot? _loaded;

    public KnownIssueCatalogService(
        HttpClient? httpClient = null,
        ICatalogSignatureVerifier? signatureVerifier = null,
        Uri? catalogUri = null,
        string? cacheDirectory = null)
    {
        _httpClient = httpClient ?? new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        _signatureVerifier = signatureVerifier ?? new EcdsaCatalogSignatureVerifier(PublicKeyPem);
        _catalogUri = catalogUri ?? new Uri("https://github.com/OneDeadMachine-Dev/W-FIX/releases/latest/download/known-issues.json");
        _cacheDirectory = cacheDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "W-Fix",
            "Catalog");
    }

    public async Task<KnownIssueCatalogSnapshot> LoadAsync(
        bool forceRefresh = false,
        CancellationToken cancellationToken = default)
    {
        if (!forceRefresh && _loaded is not null && DateTimeOffset.UtcNow - _loaded.LoadedAt < TimeSpan.FromHours(24))
            return _loaded;

        string? warning = null;
        try
        {
            var remote = await DownloadAndValidateAsync(cancellationToken);
            Directory.CreateDirectory(_cacheDirectory);
            await File.WriteAllBytesAsync(Path.Combine(_cacheDirectory, "known-issues.json"), remote.Catalog, cancellationToken);
            await File.WriteAllBytesAsync(Path.Combine(_cacheDirectory, "known-issues.sig"), remote.Signature, cancellationToken);
            _loaded = ToSnapshot(remote.Document, _catalogUri.ToString(), isFallback: false);
            return _loaded;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            warning = $"Онлайн-каталог недоступен или отклонён: {ex.Message}";
        }

        try
        {
            var cached = await LoadSignedCacheAsync(cancellationToken);
            if (cached.ExpiresAt > DateTimeOffset.UtcNow)
            {
                _loaded = ToSnapshot(cached, "signed-cache", isFallback: true, warning);
                return _loaded;
            }
            warning += " Подписанный кэш устарел.";
        }
        catch (FileNotFoundException)
        {
            // Первый запуск: подписанного кэша ещё нет.
        }
        catch (DirectoryNotFoundException)
        {
            // Первый запуск: каталог кэша ещё не создан.
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            warning += $" Кэш отклонён: {ex.Message}";
        }

        var embedded = await LoadEmbeddedAsync(cancellationToken);
        _loaded = ToSnapshot(embedded, "embedded", isFallback: true, warning);
        return _loaded;
    }

    private async Task<(KnownIssueCatalogDocument Document, byte[] Catalog, byte[] Signature)> DownloadAndValidateAsync(
        CancellationToken cancellationToken)
    {
        var signatureUri = new Uri(_catalogUri + ".sig");
        var catalog = await _httpClient.GetByteArrayAsync(_catalogUri, cancellationToken);
        var signatureText = await _httpClient.GetStringAsync(signatureUri, cancellationToken);
        var signature = Convert.FromBase64String(signatureText.Trim());
        if (!_signatureVerifier.Verify(catalog, signature))
            throw new CryptographicException("Некорректная подпись каталога.");

        var document = DeserializeAndValidate(catalog);
        if (document.ExpiresAt <= DateTimeOffset.UtcNow)
            throw new InvalidDataException("Каталог просрочен.");
        return (document, catalog, signature);
    }

    private async Task<KnownIssueCatalogDocument> LoadSignedCacheAsync(CancellationToken cancellationToken)
    {
        var catalog = await File.ReadAllBytesAsync(Path.Combine(_cacheDirectory, "known-issues.json"), cancellationToken);
        var signature = await File.ReadAllBytesAsync(Path.Combine(_cacheDirectory, "known-issues.sig"), cancellationToken);
        if (!_signatureVerifier.Verify(catalog, signature))
            throw new CryptographicException("Некорректная подпись кэша.");
        return DeserializeAndValidate(catalog);
    }

    private static async Task<KnownIssueCatalogDocument> LoadEmbeddedAsync(CancellationToken cancellationToken)
    {
        await using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("WFix.Core.Catalog.known-issues.json")
                                 ?? throw new InvalidOperationException("Встроенный каталог не найден.");
        using var memory = new MemoryStream();
        await stream.CopyToAsync(memory, cancellationToken);
        return DeserializeAndValidate(memory.ToArray());
    }

    private static KnownIssueCatalogDocument DeserializeAndValidate(byte[] data)
    {
        var document = JsonSerializer.Deserialize<KnownIssueCatalogDocument>(data, JsonOptions)
                       ?? throw new InvalidDataException("Каталог не содержит документа.");
        Validate(document);
        return document;
    }

    private static void Validate(KnownIssueCatalogDocument document)
    {
        if (document.SchemaVersion != 1)
            throw new InvalidDataException($"Неподдерживаемая схема каталога: {document.SchemaVersion}.");
        if (document.ExpiresAt <= document.GeneratedAt)
            throw new InvalidDataException("Некорректный срок действия каталога.");
        if (document.Entries.GroupBy(entry => entry.Id, StringComparer.OrdinalIgnoreCase).Any(group => group.Count() > 1))
            throw new InvalidDataException("Каталог содержит повторяющиеся идентификаторы.");

        foreach (var entry in document.Entries)
        {
            if (string.IsNullOrWhiteSpace(entry.Id) || string.IsNullOrWhiteSpace(entry.Title))
                throw new InvalidDataException("Запись каталога не имеет обязательного идентификатора или заголовка.");
            if (entry.OfficialSource.Scheme != Uri.UriSchemeHttps || !AllowedSourceHosts.Contains(entry.OfficialSource.Host))
                throw new InvalidDataException($"Источник '{entry.OfficialSource}' не является разрешённым Microsoft-источником.");
            if (entry.RecommendedActionIds.Any(action => !action.StartsWith("legacy:", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidDataException($"Запись '{entry.Id}' ссылается на неизвестное пространство действий.");
        }
    }

    private static KnownIssueCatalogSnapshot ToSnapshot(
        KnownIssueCatalogDocument document,
        string source,
        bool isFallback,
        string? warning = null) =>
        new()
        {
            Source = source,
            ExpiresAt = document.ExpiresAt,
            IsFallback = isFallback,
            Warning = warning,
            Entries = document.Entries
        };

    public void Dispose()
    {
        _httpClient.Dispose();
        if (_signatureVerifier is IDisposable disposable)
            disposable.Dispose();
    }
}
