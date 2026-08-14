using System.Text.Json.Serialization;

namespace WFix.Core.Models;

public enum TargetSource
{
    Manual,
    ActiveDirectory,
    Local
}

public sealed record TargetDescriptor
{
    public required string Id { get; init; }
    public required string ComputerName { get; init; }
    public string? Fqdn { get; init; }
    public string? OuPath { get; init; }
    public TargetSource Source { get; init; }
    public CredentialReference? Credential { get; init; }

    [JsonIgnore]
    public string ConnectionName => string.IsNullOrWhiteSpace(Fqdn) ? ComputerName : Fqdn;

    public static TargetDescriptor Local() => new()
    {
        Id = Environment.MachineName.ToUpperInvariant(),
        ComputerName = Environment.MachineName,
        Source = TargetSource.Local
    };
}

/// <summary>
/// Непрозрачная ссылка на запись Windows Credential Manager. Секрет не является частью модели задания.
/// </summary>
public sealed record CredentialReference(string TargetName, string UserName);

public sealed record RemoteCapabilityReport
{
    public required TargetDescriptor Target { get; init; }
    public DateTimeOffset CheckedAt { get; init; } = DateTimeOffset.UtcNow;
    public bool DnsResolved { get; init; }
    public bool PingResponded { get; init; }
    public long? PingMilliseconds { get; init; }
    public bool WinRmAvailable { get; init; }
    public bool CimAvailable { get; init; }
    public bool PowerShellAvailable { get; init; }
    public bool IsAdministrator { get; init; }
    public bool TaskSchedulerAvailable { get; init; }
    public bool SpoolerAvailable { get; init; }
    public bool HasInteractiveUser { get; init; }
    public string? InteractiveUser { get; init; }
    public bool PendingReboot { get; init; }
    public string? OperatingSystem { get; init; }
    public string? OperatingSystemVersion { get; init; }
    public string? Architecture { get; init; }
    public long? FreeSystemDriveBytes { get; init; }
    public IReadOnlyList<string> Errors { get; init; } = [];

    public bool CanInspect => WinRmAvailable && PowerShellAvailable;
    public bool CanRepairMachine => CanInspect && IsAdministrator && SpoolerAvailable;
    public bool CanRepairInteractiveUser => CanRepairMachine && TaskSchedulerAvailable && HasInteractiveUser;
}

public sealed record PrinterSnapshot
{
    public required string Name { get; init; }
    public string DriverName { get; init; } = "";
    public string DriverInfName { get; init; } = "";
    public string DriverManufacturer { get; init; } = "";
    public string DriverVersion { get; init; } = "";
    public string PortName { get; init; } = "";
    public string PortKind { get; init; } = "Unknown";
    public string? Address { get; init; }
    public PrinterStatus Status { get; init; }
    public bool IsDefault { get; init; }
    public bool IsShared { get; init; }
    public int JobCount { get; init; }
    public IReadOnlyList<string> ErrorCodes { get; init; } = [];
}

public sealed record PrinterDriverSnapshot
{
    public required string Name { get; init; }
    public string Manufacturer { get; init; } = "";
    public string Version { get; init; } = "";
    public string InfName { get; init; } = "";
    public string Architecture { get; init; } = "";
}

public sealed record PrintJobSnapshot
{
    public int Id { get; init; }
    public string PrinterName { get; init; } = "";
    public string Owner { get; init; } = "";
    public string Status { get; init; } = "";
    public long Size { get; init; }
}

public sealed record PrinterInventory
{
    public required TargetDescriptor Target { get; init; }
    public DateTimeOffset CapturedAt { get; init; } = DateTimeOffset.UtcNow;
    public string OperatingSystem { get; init; } = "";
    public string OperatingSystemVersion { get; init; } = "";
    public int BuildNumber { get; init; }
    public int UpdateBuildRevision { get; init; }
    public string Architecture { get; init; } = "";
    public string SpoolerStatus { get; init; } = "Unknown";
    public bool PendingReboot { get; init; }
    public bool WindowsProtectedPrintModeEnabled { get; init; }
    public bool WindowsProtectedPrintModeManaged { get; init; }
    public IReadOnlyList<string> InstalledKnowledgeBaseIds { get; init; } = [];
    public IReadOnlyList<PrinterSnapshot> Printers { get; init; } = [];
    public IReadOnlyList<PrinterDriverSnapshot> Drivers { get; init; } = [];
    public IReadOnlyList<PrintJobSnapshot> Jobs { get; init; } = [];
    public IReadOnlyDictionary<string, string?> Policies { get; init; } = new Dictionary<string, string?>();
    public IReadOnlyList<string> RecentPrintServiceErrors { get; init; } = [];
}

public enum FindingSeverity
{
    Information,
    Warning,
    Error,
    Critical
}

public sealed record DiagnosticFinding
{
    public required string RuleId { get; init; }
    public required TargetDescriptor Target { get; init; }
    public string? PrinterName { get; init; }
    public required string Title { get; init; }
    public required string Description { get; init; }
    public FindingSeverity Severity { get; init; }
    public double Confidence { get; init; }
    public IReadOnlyList<string> Evidence { get; init; } = [];
    public IReadOnlyList<string> RecommendedActionIds { get; init; } = [];
    public Uri? OfficialSource { get; init; }
    public bool VerifyResolution { get; init; } = true;
}

[Flags]
public enum RepairRequirement
{
    None = 0,
    Administrator = 1,
    WinRm = 2,
    Spooler = 4,
    TaskScheduler = 8,
    InteractiveUser = 16,
    Internet = 32
}

public enum RepairRisk
{
    ReadOnly,
    Reversible,
    Disruptive,
    Irreversible
}

public sealed record RepairStep
{
    public required string Id { get; init; }
    public required string ActionId { get; init; }
    public required string Title { get; init; }
    public string Description { get; init; } = "";
    public string? PrinterName { get; init; }
    public RepairRisk Risk { get; init; }
    public RepairRequirement Requirements { get; init; }
    public IReadOnlyList<string> DependsOn { get; init; } = [];
    public bool RequiresAdditionalConfirmation => Risk == RepairRisk.Irreversible;
}

public sealed record RepairPlan
{
    public required string Id { get; init; }
    public required TargetDescriptor Target { get; init; }
    public DateTimeOffset CreatedAt { get; init; } = DateTimeOffset.UtcNow;
    public IReadOnlyList<DiagnosticFinding> Findings { get; init; } = [];
    public IReadOnlyList<RepairStep> Steps { get; init; } = [];
    public bool RequiresReboot { get; init; }
    public bool RequiresAdditionalConfirmation => Steps.Any(step => step.RequiresAdditionalConfirmation);
}

public sealed record RepairStepResult
{
    public required string StepId { get; init; }
    public required string ActionId { get; init; }
    public bool Succeeded { get; init; }
    public bool Verified { get; init; }
    public bool RolledBack { get; init; }
    public string Summary { get; init; } = "";
    public string? SnapshotPath { get; init; }
    public IReadOnlyList<string> Output { get; init; } = [];
}

public enum RepairRunStatus
{
    Pending,
    Running,
    Succeeded,
    Failed,
    RolledBack,
    Cancelled,
    Skipped
}

public sealed record RepairRun
{
    public required string Id { get; init; }
    public required TargetDescriptor Target { get; init; }
    public DateTimeOffset StartedAt { get; init; }
    public DateTimeOffset CompletedAt { get; init; }
    public RepairRunStatus Status { get; init; }
    public bool PendingReboot { get; init; }
    public IReadOnlyList<DiagnosticFinding> Findings { get; init; } = [];
    public IReadOnlyList<RepairStepResult> Steps { get; init; } = [];
    public IReadOnlyList<string> Warnings { get; init; } = [];
    public string? ReportDirectory { get; init; }
}
