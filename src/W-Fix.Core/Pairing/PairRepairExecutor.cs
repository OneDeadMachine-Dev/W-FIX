using System.Text.Json;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class PairRepairExecutor(
    IPairActionDispatcher dispatcher,
    IPairRunReportService reportService,
    string? runRootDirectory = null) : IPairRepairExecutor
{
    private static readonly JsonSerializerOptions JournalOptions = new(JsonSerializerDefaults.Web) { WriteIndented = true };
    private readonly string _runRootDirectory = runRootDirectory ?? Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "W-Fix", "Runs");

    public async Task<PairRun> ExecuteAsync(
        PairRepairPlan plan,
        IReadOnlyDictionary<PairEndpointRole, TargetDescriptor> targets,
        IProgress<PairRun>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(targets);
        ValidatePlan(plan, targets);
        var startedAt = DateTimeOffset.UtcNow;
        var results = new List<PairStepResult>();
        var completed = new Stack<CompletedStep>();
        var warnings = new List<string>();
        var pendingReboot = false;
        var runDirectory = GetRunDirectory(plan.Id);
        Directory.CreateDirectory(runDirectory);
        var journalPath = Path.Combine(runDirectory, "pair-run.pending.json");

        Report(PairRunStatus.Running);
        try
        {
            foreach (var step in plan.Steps)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (step.DependsOn.Any(dependency => !results.Any(result => result.StepId == dependency && result.Verified)))
                    throw new PairExecutionException($"Не выполнена зависимость шага '{step.Id}'.");
                var target = targets[step.Endpoint];
                var context = new PairActionContext(target, step, runDirectory);
                var checkpoint = await dispatcher.PrepareAsync(context, cancellationToken);
                var completedStep = new CompletedStep(context, checkpoint);
                completed.Push(completedStep);
                await WriteJournalAsync(journalPath, plan, completed, startedAt, cancellationToken);

                var execution = await dispatcher.ExecuteAsync(context, cancellationToken);
                var verified = execution.Success && await dispatcher.VerifyAsync(context, cancellationToken);
                pendingReboot |= execution.RequiresReboot;
                results.Add(new PairStepResult
                {
                    StepId = step.Id,
                    ActionId = step.ActionId,
                    Endpoint = step.Endpoint,
                    Succeeded = execution.Success,
                    Verified = verified,
                    Summary = execution.Summary,
                    Output = execution.Output
                });
                if (!verified)
                    throw new PairExecutionException($"Проверка шага '{step.Title}' не подтвердила исправление.");
                Report(PairRunStatus.Running);
            }

            await dispatcher.CompleteAsync(true, cancellationToken);
            var succeeded = Build(PairRunStatus.Succeeded);
            var persisted = await PersistAsync(succeeded, cancellationToken);
            DeleteJournal(journalPath);
            progress?.Report(persisted);
            return persisted;
        }
        catch (OperationCanceledException)
        {
            var rolledBack = await RollbackAsync(completed, results, warnings);
            var cancelled = Build(rolledBack ? PairRunStatus.Cancelled : PairRunStatus.RecoveryRequired);
            var persisted = await PersistAsync(cancelled, CancellationToken.None);
            if (rolledBack) DeleteJournal(journalPath);
            progress?.Report(persisted);
            return persisted;
        }
        catch (Exception ex)
        {
            warnings.Add(ex.Message);
            var rolledBack = await RollbackAsync(completed, results, warnings);
            var failed = Build(rolledBack ? PairRunStatus.RolledBack : PairRunStatus.RecoveryRequired);
            var persisted = await PersistAsync(failed, CancellationToken.None);
            if (rolledBack) DeleteJournal(journalPath);
            progress?.Report(persisted);
            return persisted;
        }

        PairRun Build(PairRunStatus status) => new()
        {
            Id = plan.Id,
            Host = plan.Host,
            Client = plan.Client,
            TransportMode = plan.TransportMode,
            Status = status,
            StartedAt = startedAt,
            CompletedAt = DateTimeOffset.UtcNow,
            Findings = plan.Findings,
            Steps = results.ToArray(),
            Warnings = warnings.ToArray(),
            PendingReboot = pendingReboot,
            ReportDirectory = runDirectory
        };

        void Report(PairRunStatus status) => progress?.Report(Build(status));
    }

    private async Task<bool> RollbackAsync(
        Stack<CompletedStep> completed,
        List<PairStepResult> results,
        List<string> warnings)
    {
        var allSucceeded = true;
        while (completed.TryPop(out var item))
        {
            try
            {
                var result = await dispatcher.RollbackAsync(item.Context, item.Checkpoint, CancellationToken.None);
                var index = results.FindIndex(step => step.StepId == item.Context.Step.Id);
                if (index >= 0)
                    results[index] = results[index] with { RolledBack = result.Success, Output = results[index].Output.Concat(result.Output).ToArray() };
                if (!result.Success)
                {
                    allSucceeded = false;
                    warnings.Add($"Rollback '{item.Context.Step.Title}': {result.Summary}");
                }
            }
            catch (Exception ex)
            {
                allSucceeded = false;
                warnings.Add($"Rollback '{item.Context.Step.Title}': {ex.Message}");
            }
        }
        try
        {
            await dispatcher.CompleteAsync(false, CancellationToken.None);
        }
        catch (Exception ex)
        {
            allSucceeded = false;
            warnings.Add($"Pair abort: {ex.Message}");
        }
        return allSucceeded;
    }

    private async Task<PairRun> PersistAsync(PairRun run, CancellationToken cancellationToken)
    {
        try
        {
            var directory = await reportService.WriteAsync(run, cancellationToken);
            return run with { ReportDirectory = directory };
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return run with { Warnings = run.Warnings.Concat([$"Не удалось сохранить PairRun: {ex.Message}"]).ToArray() };
        }
    }

    private static void ValidatePlan(PairRepairPlan plan, IReadOnlyDictionary<PairEndpointRole, TargetDescriptor> targets)
    {
        if (!Guid.TryParseExact(plan.Id, "N", out _))
            throw new InvalidDataException("Некорректный PairRun ID.");
        if (!targets.ContainsKey(PairEndpointRole.Host) || !targets.ContainsKey(PairEndpointRole.Client))
            throw new ArgumentException("Для live/domain PairRun требуются цели Host и Client.", nameof(targets));
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var step in plan.Steps)
        {
            if (!ids.Add(step.Id)) throw new InvalidDataException($"Дублирующийся Pair step ID: {step.Id}");
            if (!step.ActionId.StartsWith("pair.", StringComparison.Ordinal)) throw new InvalidDataException("План содержит действие вне allowlist pair.*.");
            if (step.DependsOn.Any(dependency => !ids.Contains(dependency))) throw new InvalidDataException($"Шаг '{step.Id}' ссылается на неизвестную или последующую зависимость.");
        }
    }

    private string GetRunDirectory(string runId) => Path.Combine(_runRootDirectory, "pair-" + runId);

    private static async Task WriteJournalAsync(
        string path,
        PairRepairPlan plan,
        IEnumerable<CompletedStep> completed,
        DateTimeOffset startedAt,
        CancellationToken cancellationToken)
    {
        var journal = new RecoveryJournal(
            plan.Id,
            startedAt,
            plan.Host,
            plan.Client,
            completed.Reverse().Select(item => new RecoveryEntry(
                item.Context.Target with { Credential = null },
                item.Context.Step,
                item.Checkpoint.SnapshotPath,
                item.Checkpoint.State)).ToArray());
        await using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true);
        await JsonSerializer.SerializeAsync(stream, journal, JournalOptions, cancellationToken);
    }

    private static void DeleteJournal(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }

    private sealed record CompletedStep(PairActionContext Context, PairActionCheckpoint Checkpoint);
    private sealed record RecoveryJournal(string RunId, DateTimeOffset StartedAt, PairEndpointDescriptor Host, PairEndpointDescriptor Client, IReadOnlyList<RecoveryEntry> Entries);
    private sealed record RecoveryEntry(TargetDescriptor Target, PairRepairStep Step, string? SnapshotPath, IReadOnlyDictionary<string, string?> State);
    private sealed class PairExecutionException(string message) : Exception(message);
}
