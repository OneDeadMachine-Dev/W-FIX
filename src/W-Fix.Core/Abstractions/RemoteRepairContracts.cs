using System.Net;
using System.Text.Json.Serialization;
using WFix.Core.Models;

namespace WFix.Core.Abstractions;

public sealed record RemoteCommandResult(
    bool Success,
    IReadOnlyList<string> Output,
    string? Error = null,
    bool TimedOut = false);

public interface IRemoteSession : IAsyncDisposable
{
    TargetDescriptor Target { get; }
    Task<RemoteCommandResult> ExecutePowerShellAsync(
        string script,
        CancellationToken cancellationToken = default,
        TimeSpan? timeout = null);
}

public interface IRemoteSessionFactory
{
    Task<IRemoteSession> CreateAsync(TargetDescriptor target, CancellationToken cancellationToken = default);
}

public interface ICredentialStore
{
    Task SaveAsync(CredentialReference reference, NetworkCredential credential, CancellationToken cancellationToken = default);
    Task<NetworkCredential?> ReadAsync(CredentialReference reference, CancellationToken cancellationToken = default);
    Task DeleteAsync(CredentialReference reference, CancellationToken cancellationToken = default);
}

public interface IRemotePreflightService
{
    Task<RemoteCapabilityReport> CheckAsync(TargetDescriptor target, CancellationToken cancellationToken = default);
}

public interface IPrinterInventoryService
{
    Task<PrinterInventory> CaptureAsync(TargetDescriptor target, CancellationToken cancellationToken = default);
}

public interface IDiagnosticRule
{
    string Id { get; }
    Task<IReadOnlyList<DiagnosticFinding>> EvaluateAsync(
        PrinterInventory inventory,
        CancellationToken cancellationToken = default);
}

public sealed record RepairActionContext(
    TargetDescriptor Target,
    RepairStep Step,
    PrinterInfo? Printer,
    IProgress<LogEntry>? Progress);

public sealed record RepairActionCheckpoint(
    string? SnapshotPath,
    [property: JsonIgnore] object? State = null);

public sealed record RepairActionResult(
    bool Success,
    string Summary,
    IReadOnlyList<string> Output,
    bool RequiresReboot = false);

public interface IRepairAction
{
    string Id { get; }
    string Name { get; }
    string Description { get; }
    RepairRisk Risk { get; }
    RepairRequirement Requirements { get; }
    bool IsIdempotent { get; }

    Task<RepairActionCheckpoint> PrepareAsync(RepairActionContext context, CancellationToken cancellationToken = default);
    Task<RepairActionResult> ExecuteAsync(RepairActionContext context, CancellationToken cancellationToken = default);
    Task<bool> VerifyAsync(RepairActionContext context, RepairActionResult result, CancellationToken cancellationToken = default);
    Task<RepairActionResult> RollbackAsync(RepairActionContext context, RepairActionCheckpoint checkpoint, CancellationToken cancellationToken = default);
}

public interface IRepairActionRegistry
{
    IReadOnlyList<IRepairAction> GetAll();
    IRepairAction? Get(string actionId);
}

public interface IRepairPlanner
{
    RepairPlan CreatePlan(
        TargetDescriptor target,
        IReadOnlyList<DiagnosticFinding> findings,
        RemoteCapabilityReport capabilities);
}

public sealed record RepairBatchOptions
{
    public int MaxConcurrency { get; init; } = 3;
    public bool RollbackOnVerificationFailure { get; init; } = true;
}

public interface IRepairExecutor
{
    Task<IReadOnlyList<RepairRun>> ExecuteBatchAsync(
        IReadOnlyList<RepairPlan> plans,
        RepairBatchOptions options,
        IProgress<RepairRun>? progress = null,
        CancellationToken cancellationToken = default);
}

public interface IRunReportService
{
    Task<string> WriteAsync(RepairRun run, CancellationToken cancellationToken = default);
}

public interface IRemoteMaintenanceService
{
    Task<RemoteCommandResult> RestartAsync(TargetDescriptor target, CancellationToken cancellationToken = default);
}

public interface IKnownIssueCatalog
{
    Task<KnownIssueCatalogSnapshot> LoadAsync(bool forceRefresh = false, CancellationToken cancellationToken = default);
}

public interface ISupportBundleService
{
    Task<string> ExportAsync(
        IReadOnlyList<RepairRun> runs,
        string outputPath,
        CancellationToken cancellationToken = default);
}
