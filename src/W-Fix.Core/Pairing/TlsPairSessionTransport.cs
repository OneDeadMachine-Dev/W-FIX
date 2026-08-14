using System.Buffers.Binary;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class TlsPairSessionTransport(IPairInvitationValidator validator) : IPairSessionTransport
{
    public Task<IPairHost> StartHostAsync(PairHostOptions options, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(options);
        cancellationToken.ThrowIfCancellationRequested();
        if (options.InvitationLifetime < TimeSpan.FromMinutes(1) || options.InvitationLifetime > TimeSpan.FromMinutes(30))
            throw new ArgumentOutOfRangeException(nameof(options), "Приглашение должно действовать от 1 до 30 минут.");

        var advertised = options.ListenAddresses?.Distinct().ToArray() ?? GetAdvertisedAddresses();
        if (advertised.Length == 0)
            throw new InvalidOperationException("Не найден локальный IPv4-адрес для pairing-сессии.");
        var bindAddress = advertised.All(IPAddress.IsLoopback) ? IPAddress.Loopback : IPAddress.Any;
        var listener = new TcpListener(bindAddress, 0);
        listener.Start(1);
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;
        var certificate = CreateEphemeralCertificate(options.HostComputerName, options.InvitationLifetime);
        var publicKeyHash = GetPublicKeyHash(certificate);
        var createdAt = DateTimeOffset.UtcNow;
        var sessionId = Guid.NewGuid().ToString("N");
        var confirmationCode = CreateConfirmationCode(sessionId, publicKeyHash);
        var invitation = new PairInvitation
        {
            SessionId = sessionId,
            HostComputerName = options.HostComputerName,
            HostAddresses = advertised.Select(address => address.ToString()).ToArray(),
            Port = port,
            CertificatePublicKeySha256 = publicKeyHash,
            ConfirmationCode = confirmationCode,
            CreatedAt = createdAt,
            ExpiresAt = createdAt + options.InvitationLifetime,
            PrinterName = options.PrinterName,
            ShareName = options.ShareName,
            ExpectedClientComputerName = options.ExpectedClientComputerName
        };
        validator.Validate(invitation, createdAt);
        return Task.FromResult<IPairHost>(new PairHost(listener, certificate, invitation));
    }

    public async Task<IPairSession> JoinAsync(PairInvitation invitation, CancellationToken cancellationToken = default)
    {
        validator.Validate(invitation, DateTimeOffset.UtcNow);
        if (!string.IsNullOrWhiteSpace(invitation.ExpectedClientComputerName) &&
            !string.Equals(invitation.ExpectedClientComputerName, Environment.MachineName, StringComparison.OrdinalIgnoreCase))
            throw new AuthenticationException($"Приглашение предназначено для ПК '{invitation.ExpectedClientComputerName}', а не для '{Environment.MachineName}'.");
        Exception? lastError = null;
        foreach (var addressText in invitation.HostAddresses)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var client = new TcpClient(AddressFamily.InterNetwork);
            try
            {
                await client.ConnectAsync(IPAddress.Parse(addressText), invitation.Port, cancellationToken);
                var ssl = new SslStream(client.GetStream(), false);
                await ssl.AuthenticateAsClientAsync(new SslClientAuthenticationOptions
                {
                    TargetHost = invitation.HostComputerName,
                    EnabledSslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                    CertificateRevocationCheckMode = X509RevocationMode.NoCheck,
                    RemoteCertificateValidationCallback = (_, certificate, _, _) =>
                        certificate is not null && FixedTimeEqualsHex(GetPublicKeyHash(new X509Certificate2(certificate)), invitation.CertificatePublicKeySha256)
                }, cancellationToken);
                var session = new TlsPairSession(client, ssl, invitation, PairEndpointRole.Client);
                await session.SendHandshakeAsync(new PairHello(invitation.SessionId, Environment.MachineName), cancellationToken);
                var hostHello = await session.ReceiveHandshakeAsync(cancellationToken);
                if (!string.Equals(hostHello.SessionId, invitation.SessionId, StringComparison.Ordinal))
                    throw new AuthenticationException("Pair Session ID хоста не совпадает с приглашением.");
                if (!string.Equals(hostHello.ComputerName, invitation.HostComputerName, StringComparison.OrdinalIgnoreCase))
                    throw new AuthenticationException("Имя pairing-хоста не совпадает с приглашением.");
                session.SetPeerComputerName(hostHello.ComputerName);
                session.MarkAwaitingApproval();
                return session;
            }
            catch (OperationCanceledException)
            {
                client.Dispose();
                throw;
            }
            catch (Exception ex)
            {
                lastError = ex;
                client.Dispose();
            }
        }
        throw new IOException("Не удалось подключиться ни к одному адресу pairing-хоста.", lastError);
    }

    private static IPAddress[] GetAdvertisedAddresses() =>
        NetworkInterface.GetAllNetworkInterfaces()
            .Where(network => network.OperationalStatus == OperationalStatus.Up && network.NetworkInterfaceType != NetworkInterfaceType.Loopback)
            .SelectMany(network => network.GetIPProperties().UnicastAddresses)
            .Select(address => address.Address)
            .Where(address => address.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(address) && !address.ToString().StartsWith("169.254.", StringComparison.Ordinal))
            .Distinct()
            .ToArray();

    private static X509Certificate2 CreateEphemeralCertificate(string hostName, TimeSpan lifetime)
    {
        using var key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var request = new CertificateRequest($"CN={hostName}", key, HashAlgorithmName.SHA256);
        request.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, true));
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection { new("1.3.6.1.5.5.7.3.1") }, true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, false));
        var san = new SubjectAlternativeNameBuilder();
        san.AddDnsName(hostName);
        request.CertificateExtensions.Add(san.Build());
        using var created = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow + lifetime + TimeSpan.FromMinutes(2));
        // Windows Schannel cannot use an ephemeral CNG private key for server authentication.
        // Re-import without PersistKeySet: the temporary key container is removed when the certificate is disposed.
        var password = Convert.ToHexString(RandomNumberGenerator.GetBytes(24));
        var pfx = created.Export(X509ContentType.Pkcs12, password);
        try
        {
            return new X509Certificate2(pfx, password, X509KeyStorageFlags.UserKeySet | X509KeyStorageFlags.Exportable);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(pfx);
        }
    }

    private static string GetPublicKeyHash(X509Certificate2 certificate)
    {
        using var key = certificate.GetECDsaPublicKey()
                        ?? throw new AuthenticationException("Pairing certificate must use ECDSA.");
        return Convert.ToHexString(SHA256.HashData(key.ExportSubjectPublicKeyInfo())).ToLowerInvariant();
    }

    private static string CreateConfirmationCode(string sessionId, string publicKeyHash)
    {
        var hash = SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(sessionId + publicKeyHash));
        var value = BinaryPrimitives.ReadUInt32BigEndian(hash) % 1_000_000;
        return value.ToString("D6");
    }

    private static bool FixedTimeEqualsHex(string actual, string expected)
    {
        try
        {
            return CryptographicOperations.FixedTimeEquals(Convert.FromHexString(actual), Convert.FromHexString(expected));
        }
        catch (FormatException)
        {
            return false;
        }
    }

    private sealed class PairHost(TcpListener listener, X509Certificate2 certificate, PairInvitation invitation) : IPairHost
    {
        private int _accepted;

        public PairInvitation Invitation { get; } = invitation;

        public async Task<IPairSession> AcceptAsync(CancellationToken cancellationToken = default)
        {
            if (Interlocked.Exchange(ref _accepted, 1) != 0)
                throw new InvalidOperationException("Pairing-приглашение уже было использовано.");
            var remaining = Invitation.ExpiresAt - DateTimeOffset.UtcNow;
            if (remaining <= TimeSpan.Zero)
                throw new InvalidDataException("Срок действия pairing-приглашения истёк.");
            using var expiry = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            expiry.CancelAfter(remaining);
            TcpClient? client = null;
            try
            {
                client = await listener.AcceptTcpClientAsync(expiry.Token);
                listener.Stop();
                var ssl = new SslStream(client.GetStream(), false);
                await ssl.AuthenticateAsServerAsync(new SslServerAuthenticationOptions
                {
                    ServerCertificate = certificate,
                    ClientCertificateRequired = false,
                    EnabledSslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                    CertificateRevocationCheckMode = X509RevocationMode.NoCheck
                }, expiry.Token);
                var session = new TlsPairSession(client, ssl, Invitation, PairEndpointRole.Host);
                var clientHello = await session.ReceiveHandshakeAsync(expiry.Token);
                if (!string.Equals(clientHello.SessionId, Invitation.SessionId, StringComparison.Ordinal))
                    throw new AuthenticationException("Pair Session ID клиента не совпадает с приглашением.");
                if (!string.IsNullOrWhiteSpace(Invitation.ExpectedClientComputerName) &&
                    !string.Equals(clientHello.ComputerName, Invitation.ExpectedClientComputerName, StringComparison.OrdinalIgnoreCase))
                    throw new AuthenticationException("К pairing-хосту подключился другой компьютер.");
                session.SetPeerComputerName(clientHello.ComputerName);
                await session.SendHandshakeAsync(new PairHello(Invitation.SessionId, Invitation.HostComputerName), expiry.Token);
                session.MarkAwaitingApproval();
                return session;
            }
            catch
            {
                client?.Dispose();
                throw;
            }
        }

        public ValueTask DisposeAsync()
        {
            listener.Stop();
            certificate.Dispose();
            return ValueTask.CompletedTask;
        }
    }

    private sealed class TlsPairSession(
        TcpClient client,
        SslStream stream,
        PairInvitation invitation,
        PairEndpointRole localRole) : IPairSession
    {
        private readonly PairProtocolSerializer _serializer = new();
        private readonly SemaphoreSlim _writeGate = new(1, 1);
        private int _disposed;

        public PairInvitation Invitation { get; } = invitation;
        public PairEndpointRole LocalRole { get; } = localRole;
        public string PeerComputerName { get; private set; } = "";
        public PairSessionState State { get; private set; } = PairSessionState.Connected;
        public string ConfirmationCode => Invitation.ConfirmationCode;

        public async Task<bool> ApproveAsync(bool approved, CancellationToken cancellationToken = default)
        {
            if (State != PairSessionState.AwaitingApproval)
                throw new InvalidOperationException("Pairing-сессия не ожидает подтверждения.");
            await SendCoreAsync(PairMessageKind.Approval, new PairApproval(approved), cancellationToken);
            var peer = await ReceiveCoreAsync<PairApproval>(PairMessageKind.Approval, cancellationToken);
            var accepted = approved && peer.Approved;
            State = accepted ? PairSessionState.Approved : PairSessionState.Closed;
            return accepted;
        }

        public Task SendAsync<T>(PairMessageKind kind, T message, CancellationToken cancellationToken = default)
        {
            EnsureApproved();
            return SendCoreAsync(kind, message, cancellationToken);
        }

        public Task<T> ReceiveAsync<T>(PairMessageKind expectedKind, CancellationToken cancellationToken = default)
        {
            EnsureApproved();
            return ReceiveCoreAsync<T>(expectedKind, cancellationToken);
        }

        internal Task SendHandshakeAsync(PairHello hello, CancellationToken cancellationToken) =>
            SendCoreAsync(PairMessageKind.Hello, hello, cancellationToken);

        internal Task<PairHello> ReceiveHandshakeAsync(CancellationToken cancellationToken) =>
            ReceiveCoreAsync<PairHello>(PairMessageKind.Hello, cancellationToken);

        internal void MarkAwaitingApproval() => State = PairSessionState.AwaitingApproval;
        internal void SetPeerComputerName(string value) => PeerComputerName = value;

        private async Task SendCoreAsync<T>(PairMessageKind kind, T message, CancellationToken cancellationToken)
        {
            ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) != 0, this);
            var payload = _serializer.Serialize(kind, message);
            var header = new byte[sizeof(int)];
            BinaryPrimitives.WriteInt32BigEndian(header, payload.Length);
            await _writeGate.WaitAsync(cancellationToken);
            try
            {
                await stream.WriteAsync(header, cancellationToken);
                await stream.WriteAsync(payload, cancellationToken);
                await stream.FlushAsync(cancellationToken);
            }
            finally
            {
                _writeGate.Release();
            }
        }

        private async Task<T> ReceiveCoreAsync<T>(PairMessageKind expectedKind, CancellationToken cancellationToken)
        {
            ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) != 0, this);
            var header = new byte[sizeof(int)];
            await stream.ReadExactlyAsync(header, cancellationToken);
            var length = BinaryPrimitives.ReadInt32BigEndian(header);
            if (length is <= 0 or > PairProtocolSerializer.MaximumFrameBytes)
                throw new InvalidDataException("Некорректный размер pairing frame.");
            var payload = new byte[length];
            await stream.ReadExactlyAsync(payload, cancellationToken);
            return _serializer.Deserialize<T>(expectedKind, payload);
        }

        private void EnsureApproved()
        {
            if (State != PairSessionState.Approved)
                throw new InvalidOperationException("Pairing-сессия ещё не подтверждена обеими сторонами.");
        }

        public async ValueTask DisposeAsync()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
                return;
            State = PairSessionState.Closed;
            _writeGate.Dispose();
            await stream.DisposeAsync();
            client.Dispose();
        }
    }
}
