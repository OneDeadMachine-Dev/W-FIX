using System.Net;
using WFix.Core.Models;

namespace WFix.Core.Abstractions;

public interface IPairInvitationValidator
{
    void Validate(PairInvitation invitation, DateTimeOffset now);
}

public interface IPairFileService
{
    Task WriteInvitationAsync(string path, PairInvitation invitation, CancellationToken cancellationToken = default);
    Task<PairInvitation> ReadInvitationAsync(string path, CancellationToken cancellationToken = default);
    Task WriteOfflineSnapshotAsync(string path, PairEndpointSnapshot snapshot, CancellationToken cancellationToken = default);
    Task<PairEndpointSnapshot> ReadOfflineSnapshotAsync(string path, CancellationToken cancellationToken = default);
}

public interface IPairSession : IAsyncDisposable
{
    PairInvitation Invitation { get; }
    PairEndpointRole LocalRole { get; }
    string PeerComputerName { get; }
    PairSessionState State { get; }
    string ConfirmationCode { get; }

    Task<bool> ApproveAsync(bool approved, CancellationToken cancellationToken = default);
    Task SendAsync<T>(PairMessageKind kind, T message, CancellationToken cancellationToken = default);
    Task<T> ReceiveAsync<T>(PairMessageKind expectedKind, CancellationToken cancellationToken = default);
}

public interface IPairHost : IAsyncDisposable
{
    PairInvitation Invitation { get; }
    Task<IPairSession> AcceptAsync(CancellationToken cancellationToken = default);
}

public sealed record PairHostOptions
{
    public string HostComputerName { get; init; } = Environment.MachineName;
    public IReadOnlyList<IPAddress>? ListenAddresses { get; init; }
    public string? PrinterName { get; init; }
    public string? ShareName { get; init; }
    public string? ExpectedClientComputerName { get; init; }
    public TimeSpan InvitationLifetime { get; init; } = TimeSpan.FromMinutes(15);
}

public interface IPairSessionTransport
{
    Task<IPairHost> StartHostAsync(PairHostOptions options, CancellationToken cancellationToken = default);
    Task<IPairSession> JoinAsync(PairInvitation invitation, CancellationToken cancellationToken = default);
}

public interface IPairFirewallLeaseService
{
    Task<IAsyncDisposable> OpenAsync(string sessionId, int port, string executablePath, CancellationToken cancellationToken = default);
    Task CleanupStaleAsync(CancellationToken cancellationToken = default);
}

public interface IPairInventoryService
{
    Task<PairEndpointSnapshot> CaptureAsync(
        TargetDescriptor target,
        PairEndpointRole role,
        string peerName,
        string? printerName = null,
        string? shareName = null,
        CancellationToken cancellationToken = default);
}

public interface IPairDiagnosticRule
{
    string Id { get; }
    Task<IReadOnlyList<PairDiagnosticFinding>> EvaluateAsync(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        CancellationToken cancellationToken = default);
}

public interface IPairDiagnosticService
{
    Task<IReadOnlyList<PairDiagnosticFinding>> DiagnoseAsync(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        CancellationToken cancellationToken = default);
}

public interface IPairRepairPlanner
{
    PairRepairPlan CreatePlan(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        PairTransportMode transportMode,
        IReadOnlyList<PairDiagnosticFinding> findings,
        bool includeExpertActions = false);
}

public sealed record PairActionContext(
    TargetDescriptor Target,
    PairRepairStep Step,
    string RunDirectory,
    IProgress<LogEntry>? Progress = null);

public interface IPairRepairAction
{
    string Id { get; }
    string Name { get; }
    RepairRisk Risk { get; }
    bool ExpertOnly { get; }
    bool IsIdempotent { get; }

    Task<PairActionCheckpoint> PrepareAsync(PairActionContext context, CancellationToken cancellationToken = default);
    Task<PairActionResult> ExecuteAsync(PairActionContext context, CancellationToken cancellationToken = default);
    Task<bool> VerifyAsync(PairActionContext context, CancellationToken cancellationToken = default);
    Task<PairActionResult> RollbackAsync(PairActionContext context, PairActionCheckpoint checkpoint, CancellationToken cancellationToken = default);
}

public interface IPairRepairActionRegistry
{
    IReadOnlyList<IPairRepairAction> GetAll();
    IPairRepairAction? Get(string actionId);
}

public interface IPairActionDispatcher
{
    Task<PairActionCheckpoint> PrepareAsync(PairActionContext context, CancellationToken cancellationToken = default);
    Task<PairActionResult> ExecuteAsync(PairActionContext context, CancellationToken cancellationToken = default);
    Task<bool> VerifyAsync(PairActionContext context, CancellationToken cancellationToken = default);
    Task<PairActionResult> RollbackAsync(PairActionContext context, PairActionCheckpoint checkpoint, CancellationToken cancellationToken = default);
    Task CompleteAsync(bool commit, CancellationToken cancellationToken = default) => Task.CompletedTask;
}

public interface IPairAgentCommandLoop
{
    Task RunAsync(IPairSession session, CancellationToken cancellationToken = default);
}

public interface IPairRepairExecutor
{
    Task<PairRun> ExecuteAsync(
        PairRepairPlan plan,
        IReadOnlyDictionary<PairEndpointRole, TargetDescriptor> targets,
        IProgress<PairRun>? progress = null,
        CancellationToken cancellationToken = default);
}

public interface INetworkCredentialProvisioner
{
    Task SaveForHostAsync(string hostName, NetworkCredential credential, CancellationToken cancellationToken = default);
    Task DeleteForHostAsync(string hostName, CancellationToken cancellationToken = default);
}

public interface IPairRunReportService
{
    Task<string> WriteAsync(PairRun run, CancellationToken cancellationToken = default);
}
