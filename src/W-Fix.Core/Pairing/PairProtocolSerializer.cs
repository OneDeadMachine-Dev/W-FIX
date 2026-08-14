using System.Text.Json;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

internal sealed class PairProtocolSerializer
{
    public const int ProtocolVersion = 1;
    public const int MaximumFrameBytes = 1024 * 1024;

    private static readonly IReadOnlyDictionary<PairMessageKind, Type> AllowedTypes =
        new Dictionary<PairMessageKind, Type>
        {
            [PairMessageKind.Hello] = typeof(PairHello),
            [PairMessageKind.Approval] = typeof(PairApproval),
            [PairMessageKind.Snapshot] = typeof(PairEndpointSnapshot),
            [PairMessageKind.Plan] = typeof(PairRepairPlan),
            [PairMessageKind.ActionRequest] = typeof(PairActionRequest),
            [PairMessageKind.ActionResult] = typeof(PairActionResponse),
            [PairMessageKind.RollbackRequest] = typeof(PairControlMessage),
            [PairMessageKind.Commit] = typeof(PairControlMessage),
            [PairMessageKind.Heartbeat] = typeof(PairControlMessage),
            [PairMessageKind.Error] = typeof(PairControlMessage)
        };

    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web)
    {
        PropertyNameCaseInsensitive = false,
        MaxDepth = 32
    };
    private static readonly HashSet<string> AllowedParameterNames = new(StringComparer.Ordinal)
    {
        "hostName", "printerName", "shareName"
    };

    public byte[] Serialize<T>(PairMessageKind kind, T message)
    {
        ValidateType<T>(kind);
        ValidateMessage(message);
        var payload = JsonSerializer.SerializeToElement(message, Options);
        var bytes = JsonSerializer.SerializeToUtf8Bytes(new Envelope(ProtocolVersion, kind, payload), Options);
        if (bytes.Length > MaximumFrameBytes)
            throw new InvalidDataException("Pairing-сообщение превышает допустимый размер.");
        return bytes;
    }

    public T Deserialize<T>(PairMessageKind expectedKind, ReadOnlySpan<byte> bytes)
    {
        ValidateType<T>(expectedKind);
        if (bytes.Length is 0 or > MaximumFrameBytes)
            throw new InvalidDataException("Некорректный размер pairing-сообщения.");
        var envelope = JsonSerializer.Deserialize<Envelope>(bytes, Options)
                       ?? throw new InvalidDataException("Пустое pairing-сообщение.");
        if (envelope.Version != ProtocolVersion)
            throw new InvalidDataException($"Неподдерживаемая версия pairing-протокола: {envelope.Version}.");
        if (envelope.Kind != expectedKind)
            throw new InvalidDataException($"Ожидалось сообщение {expectedKind}, получено {envelope.Kind}.");
        var value = envelope.Payload.Deserialize<T>(Options)
                    ?? throw new InvalidDataException("Не удалось разобрать pairing-сообщение.");
        ValidateMessage(value);
        return value;
    }

    private static void ValidateType<T>(PairMessageKind kind)
    {
        if (!AllowedTypes.TryGetValue(kind, out var allowed) || allowed != typeof(T))
            throw new InvalidOperationException($"Тип {typeof(T).Name} не разрешён для сообщения {kind}.");
    }

    private static void ValidateMessage<T>(T message)
    {
        ArgumentNullException.ThrowIfNull(message);
        if (message is PairActionRequest request)
        {
            if (!Guid.TryParseExact(request.RequestId, "N", out _) || !Enum.IsDefined(request.Operation))
                throw new InvalidDataException("Некорректный pairing action request.");
            ValidateStep(request.Step);
        }
        if (message is PairRepairPlan plan)
            foreach (var step in plan.Steps) ValidateStep(step);
    }

    private static void ValidateStep(PairRepairStep step)
    {
        if (!step.ActionId.StartsWith("pair.", StringComparison.Ordinal) || step.ActionId.Length > 128)
            throw new InvalidDataException("Через pairing-транспорт разрешены только встроенные действия pair.*.");
        if (step.Parameters.Count > AllowedParameterNames.Count || step.Parameters.Any(pair =>
                !AllowedParameterNames.Contains(pair.Key) || pair.Value.Length > 1024))
            throw new InvalidDataException("Pairing-действие содержит параметр вне allowlist.");
    }

    private sealed record Envelope(int Version, PairMessageKind Kind, JsonElement Payload);
}
