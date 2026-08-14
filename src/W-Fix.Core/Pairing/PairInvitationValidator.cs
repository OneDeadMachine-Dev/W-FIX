using System.Net;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class PairInvitationValidator : IPairInvitationValidator
{
    public void Validate(PairInvitation invitation, DateTimeOffset now)
    {
        ArgumentNullException.ThrowIfNull(invitation);
        if (invitation.SchemaVersion != PairInvitation.CurrentSchemaVersion)
            throw new InvalidDataException($"Неподдерживаемая версия pairing-файла: {invitation.SchemaVersion}.");
        if (!Guid.TryParseExact(invitation.SessionId, "N", out _))
            throw new InvalidDataException("Некорректный Pair Session ID.");
        if (string.IsNullOrWhiteSpace(invitation.HostComputerName) || invitation.HostComputerName.Length > 255)
            throw new InvalidDataException("В приглашении отсутствует корректное имя хоста.");
        if (invitation.ExpectedClientComputerName is { Length: > 255 } ||
            invitation.ExpectedClientComputerName is not null && string.IsNullOrWhiteSpace(invitation.ExpectedClientComputerName))
            throw new InvalidDataException("В приглашении указано некорректное имя ожидаемого клиента.");
        if (invitation.Port is < IPEndPoint.MinPort or > IPEndPoint.MaxPort)
            throw new InvalidDataException("Некорректный TCP-порт pairing-сессии.");
        if (invitation.HostAddresses.Count is < 1 or > 16 || invitation.HostAddresses.Any(address => !IPAddress.TryParse(address, out _)))
            throw new InvalidDataException("Приглашение не содержит корректных адресов хоста.");
        if (invitation.CertificatePublicKeySha256.Length != 64 ||
            invitation.CertificatePublicKeySha256.Any(character => !Uri.IsHexDigit(character)))
            throw new InvalidDataException("Некорректный отпечаток временного TLS-ключа.");
        if (invitation.ConfirmationCode.Length != 6 || invitation.ConfirmationCode.Any(character => !char.IsAsciiDigit(character)))
            throw new InvalidDataException("Некорректный код подтверждения pairing-сессии.");
        if (invitation.ExpiresAt <= invitation.CreatedAt || invitation.ExpiresAt - invitation.CreatedAt > TimeSpan.FromMinutes(30))
            throw new InvalidDataException("Некорректный срок действия приглашения.");
        if (now < invitation.CreatedAt - TimeSpan.FromMinutes(2))
            throw new InvalidDataException("Часы компьютеров расходятся: приглашение создано в будущем.");
        if (now >= invitation.ExpiresAt)
            throw new InvalidDataException("Срок действия pairing-приглашения истёк.");
    }
}
