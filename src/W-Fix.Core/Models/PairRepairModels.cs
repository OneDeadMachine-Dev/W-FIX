using System.Text.Json.Serialization;

namespace WFix.Core.Models;

public enum PairEndpointRole
{
    Host,
    Client
}

public enum PairTransportMode
{
    DomainRemote,
    LiveLan,
    Offline
}

public enum PairSessionState
{
    Waiting,
    Connected,
    AwaitingApproval,
    Approved,
    Closed,
    Failed
}

public sealed record PairEndpointDescriptor
{
    public required PairEndpointRole Role { get; init; }
    public required string ComputerName { get; init; }
    public string? Fqdn { get; init; }
    public bool IsLocalAgent { get; init; }

    [JsonIgnore]
    public string ConnectionName => string.IsNullOrWhiteSpace(Fqdn) ? ComputerName : Fqdn;
}

public sealed record PairInvitation
{
    public const int CurrentSchemaVersion = 1;

    public int SchemaVersion { get; init; } = CurrentSchemaVersion;
    public required string SessionId { get; init; }
    public required string HostComputerName { get; init; }
    public IReadOnlyList<string> HostAddresses { get; init; } = [];
    public int Port { get; init; }
    public required string CertificatePublicKeySha256 { get; init; }
    public required string ConfirmationCode { get; init; }
    public DateTimeOffset CreatedAt { get; init; }
    public DateTimeOffset ExpiresAt { get; init; }
    public string? PrinterName { get; init; }
    public string? ShareName { get; init; }
}

public sealed record PairOfflineBundle
{
    public const int CurrentSchemaVersion = 1;

    public int SchemaVersion { get; init; } = CurrentSchemaVersion;
    public required string BundleId { get; init; }
    public DateTimeOffset CreatedAt { get; init; }
    public required string SnapshotPayloadBase64 { get; init; }
    public required string SigningPublicKeyBase64 { get; init; }
    public required string SignatureBase64 { get; init; }
}

public sealed record PairEndpointSnapshot
{
    public required PairEndpointDescriptor Endpoint { get; init; }
    public DateTimeOffset CapturedAt { get; init; } = DateTimeOffset.UtcNow;
    public string OperatingSystem { get; init; } = "";
    public string OperatingSystemVersion { get; init; } = "";
    public int BuildNumber { get; init; }
    public int UpdateBuildRevision { get; init; }
    public bool DomainJoined { get; init; }
    public string DomainOrWorkgroup { get; init; } = "";
    public string NetworkProfile { get; init; } = "Unknown";
    public IReadOnlyList<string> Ipv4Addresses { get; init; } = [];
    public bool PeerNameResolved { get; init; }
    public bool SmbPortReachable { get; init; }
    public bool RpcEndpointMapperReachable { get; init; }
    public bool SpoolerRunning { get; init; }
    public IReadOnlyDictionary<string, string> ServiceStates { get; init; } = new Dictionary<string, string>();
    public bool NetworkDiscoveryFirewallEnabled { get; init; }
    public bool FileAndPrinterSharingFirewallEnabled { get; init; }
    public bool SmbSigningRequired { get; init; }
    public bool InsecureGuestLogonsEnabled { get; init; }
    public bool HasConflictingSmbConnection { get; init; }
    public string? SmbConnectionError { get; init; }
    public bool RpcOverNamedPipes { get; init; }
    public bool RpcListenerAllowsNamedPipes { get; init; }
    public bool RpcPrivacyDisabled { get; init; }
    public bool RestrictDriverInstallationToAdministrators { get; init; } = true;
    public string? PrinterName { get; init; }
    public string? PrinterShareName { get; init; }
    public bool PrinterShared { get; init; }
    public string? PrinterDriverName { get; init; }
    public string? PrinterDriverVersion { get; init; }
    public bool PrinterConnectionInstalled { get; init; }
    public IReadOnlyList<string> RecentErrors { get; init; } = [];
}

public sealed record PairDiagnosticFinding
{
    public required string RuleId { get; init; }
    public required string Title { get; init; }
    public required string Description { get; init; }
    public FindingSeverity Severity { get; init; }
    public double Confidence { get; init; }
    public IReadOnlyList<PairEndpointRole> AffectedEndpoints { get; init; } = [];
    public IReadOnlyList<string> Evidence { get; init; } = [];
    public IReadOnlyList<string> RecommendedActionIds { get; init; } = [];
    public bool ExpertOnly { get; init; }
    public Uri? OfficialSource { get; init; }
}

public sealed record PairRepairStep
{
    public required string Id { get; init; }
    public required string ActionId { get; init; }
    public required PairEndpointRole Endpoint { get; init; }
    public required string Title { get; init; }
    public string Description { get; init; } = "";
    public RepairRisk Risk { get; init; }
    public bool ExpertOnly { get; init; }
    public IReadOnlyDictionary<string, string> Parameters { get; init; } = new Dictionary<string, string>();
    public IReadOnlyList<string> DependsOn { get; init; } = [];
    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(45);

    public bool RequiresAdditionalConfirmation => ExpertOnly || Risk is RepairRisk.Disruptive or RepairRisk.Irreversible;
}

public sealed record PairRepairPlan
{
    public required string Id { get; init; }
    public required PairEndpointDescriptor Host { get; init; }
    public required PairEndpointDescriptor Client { get; init; }
    public PairTransportMode TransportMode { get; init; }
    public DateTimeOffset CreatedAt { get; init; } = DateTimeOffset.UtcNow;
    public IReadOnlyList<PairDiagnosticFinding> Findings { get; init; } = [];
    public IReadOnlyList<PairRepairStep> Steps { get; init; } = [];

    public bool RequiresAdditionalConfirmation => Steps.Any(step => step.RequiresAdditionalConfirmation);
}

public sealed record PairActionCheckpoint
{
    public required string ActionId { get; init; }
    public required PairEndpointRole Endpoint { get; init; }
    [JsonIgnore]
    public string? SnapshotPath { get; init; }
    [JsonIgnore]
    public IReadOnlyDictionary<string, string?> State { get; init; } = new Dictionary<string, string?>();
}

public sealed record PairActionResult
{
    public bool Success { get; init; }
    public bool Verified { get; init; }
    public string Summary { get; init; } = "";
    public IReadOnlyList<string> Output { get; init; } = [];
    public bool RequiresReboot { get; init; }
}

public sealed record PairStepResult
{
    public required string StepId { get; init; }
    public required string ActionId { get; init; }
    public required PairEndpointRole Endpoint { get; init; }
    public bool Succeeded { get; init; }
    public bool Verified { get; init; }
    public bool RolledBack { get; init; }
    public string Summary { get; init; } = "";
    public IReadOnlyList<string> Output { get; init; } = [];
}

public enum PairRunStatus
{
    Pending,
    Running,
    Succeeded,
    Failed,
    RolledBack,
    Cancelled,
    RecoveryRequired
}

public sealed record PairRun
{
    public required string Id { get; init; }
    public required PairEndpointDescriptor Host { get; init; }
    public required PairEndpointDescriptor Client { get; init; }
    public PairTransportMode TransportMode { get; init; }
    public PairRunStatus Status { get; init; }
    public DateTimeOffset StartedAt { get; init; }
    public DateTimeOffset CompletedAt { get; init; }
    public IReadOnlyList<PairDiagnosticFinding> Findings { get; init; } = [];
    public IReadOnlyList<PairStepResult> Steps { get; init; } = [];
    public IReadOnlyList<string> Warnings { get; init; } = [];
    public bool PendingReboot { get; init; }
    public string? ReportDirectory { get; init; }
}

public enum PairMessageKind
{
    Hello,
    Approval,
    Snapshot,
    Plan,
    ActionRequest,
    ActionResult,
    RollbackRequest,
    Commit,
    Heartbeat,
    Error
}

public sealed record PairHello(string SessionId, string ComputerName);
public sealed record PairApproval(bool Approved);
public enum PairActionOperation
{
    Prepare,
    Execute,
    Verify,
    Rollback,
    Commit
}

public sealed record PairActionRequest(string RequestId, PairActionOperation Operation, PairRepairStep Step);
public sealed record PairActionResponse(string RequestId, PairActionResult Result);
public sealed record PairControlMessage(string RunId, string? StepId = null, string? Message = null);
