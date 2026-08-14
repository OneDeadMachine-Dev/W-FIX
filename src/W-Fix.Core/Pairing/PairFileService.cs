using System.Security.Cryptography;
using System.Text.Json;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class PairFileService(IPairInvitationValidator invitationValidator) : IPairFileService
{
    private const int MaximumFileBytes = 1024 * 1024;
    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = false,
        MaxDepth = 32
    };

    public Task WriteInvitationAsync(string path, PairInvitation invitation, CancellationToken cancellationToken = default)
    {
        invitationValidator.Validate(invitation, DateTimeOffset.UtcNow);
        return WriteAtomicAsync(path, invitation, cancellationToken);
    }

    public async Task<PairInvitation> ReadInvitationAsync(string path, CancellationToken cancellationToken = default)
    {
        var invitation = await ReadAsync<PairInvitation>(path, cancellationToken);
        invitationValidator.Validate(invitation, DateTimeOffset.UtcNow);
        return invitation;
    }

    public async Task WriteOfflineSnapshotAsync(
        string path,
        PairEndpointSnapshot snapshot,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        var payload = JsonSerializer.SerializeToUtf8Bytes(snapshot, Options);
        if (payload.Length > MaximumFileBytes / 2)
            throw new InvalidDataException("Снимок Pair Repair слишком велик для безопасного обмена.");
        using var key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var bundle = new PairOfflineBundle
        {
            BundleId = Guid.NewGuid().ToString("N"),
            CreatedAt = DateTimeOffset.UtcNow,
            SnapshotPayloadBase64 = Convert.ToBase64String(payload),
            SigningPublicKeyBase64 = Convert.ToBase64String(key.ExportSubjectPublicKeyInfo()),
            SignatureBase64 = Convert.ToBase64String(key.SignData(payload, HashAlgorithmName.SHA256))
        };
        await WriteAtomicAsync(path, bundle, cancellationToken);
    }

    public async Task<PairEndpointSnapshot> ReadOfflineSnapshotAsync(
        string path,
        CancellationToken cancellationToken = default)
    {
        var bundle = await ReadAsync<PairOfflineBundle>(path, cancellationToken);
        if (bundle.SchemaVersion != PairOfflineBundle.CurrentSchemaVersion ||
            !Guid.TryParseExact(bundle.BundleId, "N", out _) ||
            bundle.CreatedAt > DateTimeOffset.UtcNow + TimeSpan.FromMinutes(2))
            throw new InvalidDataException("Некорректный формат offline pairing bundle.");
        byte[] payload;
        byte[] publicKey;
        byte[] signature;
        try
        {
            payload = Convert.FromBase64String(bundle.SnapshotPayloadBase64);
            publicKey = Convert.FromBase64String(bundle.SigningPublicKeyBase64);
            signature = Convert.FromBase64String(bundle.SignatureBase64);
        }
        catch (FormatException ex)
        {
            throw new InvalidDataException("Offline pairing bundle содержит повреждённые данные.", ex);
        }
        if (payload.Length is 0 or > MaximumFileBytes / 2 || publicKey.Length > 1024 || signature.Length > 1024)
            throw new InvalidDataException("Offline pairing bundle превышает допустимые ограничения.");
        using var key = ECDsa.Create();
        try
        {
            key.ImportSubjectPublicKeyInfo(publicKey, out var bytesRead);
            if (bytesRead != publicKey.Length || !key.VerifyData(payload, signature, HashAlgorithmName.SHA256))
                throw new InvalidDataException("Подпись offline pairing bundle недействительна.");
        }
        catch (CryptographicException ex)
        {
            throw new InvalidDataException("Подпись offline pairing bundle недействительна.", ex);
        }
        return JsonSerializer.Deserialize<PairEndpointSnapshot>(payload, Options)
               ?? throw new InvalidDataException("Offline pairing bundle не содержит снимок.");
    }

    private static async Task<T> ReadAsync<T>(string path, CancellationToken cancellationToken)
    {
        ValidatePath(path);
        var info = new FileInfo(path);
        if (!info.Exists || info.Length is <= 0 or > MaximumFileBytes)
            throw new InvalidDataException("Pairing-файл отсутствует, пуст или превышает допустимый размер.");
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true);
        return await JsonSerializer.DeserializeAsync<T>(stream, Options, cancellationToken)
               ?? throw new InvalidDataException("Pairing-файл пуст.");
    }

    private static async Task WriteAtomicAsync<T>(string path, T value, CancellationToken cancellationToken)
    {
        ValidatePath(path);
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var temporary = fullPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, true))
            {
                await JsonSerializer.SerializeAsync(stream, value, Options, cancellationToken);
                await stream.FlushAsync(cancellationToken);
                if (stream.Length > MaximumFileBytes)
                    throw new InvalidDataException("Pairing-файл превышает допустимый размер.");
            }
            File.Move(temporary, fullPath, true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    private static void ValidatePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Путь pairing-файла не задан.", nameof(path));
    }
}
