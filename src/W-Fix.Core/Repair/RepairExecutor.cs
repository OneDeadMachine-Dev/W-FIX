using System.Collections.Concurrent;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Repair;

public sealed class RepairExecutor(
    IRepairActionRegistry registry,
    IPrinterInventoryService inventoryService,
    IEnumerable<IDiagnosticRule> diagnosticRules,
    IRunReportService reportService) : IRepairExecutor
{
    private readonly IReadOnlyDictionary<string, IDiagnosticRule> _rules = diagnosticRules
        .ToDictionary(rule => rule.Id, StringComparer.OrdinalIgnoreCase);

    public async Task<IReadOnlyList<RepairRun>> ExecuteBatchAsync(
        IReadOnlyList<RepairPlan> plans,
        RepairBatchOptions options,
        IProgress<RepairRun>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(plans);
        ArgumentNullException.ThrowIfNull(options);
        if (options.MaxConcurrency is < 1 or > 10)
            throw new ArgumentOutOfRangeException(nameof(options.MaxConcurrency), "Допустимый параллелизм: от 1 до 10.");

        using var gate = new SemaphoreSlim(options.MaxConcurrency, options.MaxConcurrency);
        var results = new ConcurrentDictionary<int, RepairRun>();
        var tasks = plans.Select((plan, index) => ExecuteWithGateAsync(plan, index)).ToArray();
        await Task.WhenAll(tasks);
        return results.OrderBy(pair => pair.Key).Select(pair => pair.Value).ToArray();

        async Task ExecuteWithGateAsync(RepairPlan plan, int index)
        {
            var entered = false;
            try
            {
                await gate.WaitAsync(cancellationToken);
                entered = true;
                results[index] = await ExecuteTargetAsync(plan, options, progress, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                results[index] = new RepairRun
                {
                    Id = plan.Id,
                    Target = plan.Target,
                    StartedAt = DateTimeOffset.UtcNow,
                    CompletedAt = DateTimeOffset.UtcNow,
                    Status = RepairRunStatus.Cancelled,
                    Findings = plan.Findings
                };
            }
            finally
            {
                if (entered)
                    gate.Release();
            }
        }
    }

    private async Task<RepairRun> ExecuteTargetAsync(
        RepairPlan plan,
        RepairBatchOptions options,
        IProgress<RepairRun>? progress,
        CancellationToken cancellationToken)
    {
        var startedAt = DateTimeOffset.UtcNow;
        var stepResults = new List<RepairStepResult>();
        var completed = new Stack<(IRepairAction Action, RepairActionContext Context, RepairActionCheckpoint Checkpoint, int ResultIndex)>();
        var warnings = new List<string>();
        PrinterInventory? inventory = null;
        var rollbackPerformed = false;

        Report(RepairRunStatus.Running);

        try
        {
            inventory = await inventoryService.CaptureAsync(plan.Target, cancellationToken);
            foreach (var step in plan.Steps)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var action = registry.Get(step.ActionId);
                if (action is null)
                {
                    stepResults.Add(Failed(step, $"Действие '{step.ActionId}' не зарегистрировано."));
                    throw new RepairExecutionException();
                }

                var printer = CreatePrinterInfo(inventory, step.PrinterName);
                var actionProgress = new Progress<LogEntry>();
                var context = new RepairActionContext(plan.Target, step, printer, actionProgress);
                RepairActionCheckpoint checkpoint;
                try
                {
                    checkpoint = await action.PrepareAsync(context, cancellationToken);
                    // Снимок кладётся в стек до мутации: даже частично упавшее действие должно быть откатано.
                    var pendingResultIndex = stepResults.Count;
                    stepResults.Add(new RepairStepResult
                    {
                        StepId = step.Id,
                        ActionId = step.ActionId,
                        Summary = "Снимок создан, действие подготовлено.",
                        SnapshotPath = checkpoint.SnapshotPath
                    });
                    completed.Push((action, context, checkpoint, pendingResultIndex));
                    var result = await action.ExecuteAsync(context, cancellationToken);
                    var verified = result.Success && await action.VerifyAsync(context, result, cancellationToken);
                    stepResults[pendingResultIndex] = new RepairStepResult
                    {
                        StepId = step.Id,
                        ActionId = step.ActionId,
                        Succeeded = result.Success,
                        Verified = verified,
                        Summary = result.Summary,
                        SnapshotPath = checkpoint.SnapshotPath,
                        Output = result.Output
                    };

                    if (!verified)
                        throw new RepairExecutionException();

                    Report(RepairRunStatus.Running);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (RepairExecutionException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    var existingIndex = stepResults.FindIndex(item => item.StepId == step.Id);
                    if (existingIndex >= 0)
                        stepResults[existingIndex] = stepResults[existingIndex] with { Summary = ex.Message };
                    else
                        stepResults.Add(Failed(step, ex.Message));
                    throw new RepairExecutionException();
                }
            }

            if (!await VerifyFindingsResolvedAsync(plan, cancellationToken))
            {
                warnings.Add("Повторная диагностика обнаружила исходную проблему после выполнения плана.");
                throw new RepairExecutionException();
            }

            var succeeded = BuildRun(RepairRunStatus.Succeeded);
            return await PersistAsync(succeeded, progress, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            if (options.RollbackOnVerificationFailure)
                rollbackPerformed = await RollbackAsync(completed, stepResults, warnings, CancellationToken.None);
            var cancelled = BuildRun(rollbackPerformed ? RepairRunStatus.RolledBack : RepairRunStatus.Cancelled);
            return await PersistAsync(cancelled, progress, CancellationToken.None);
        }
        catch (RepairExecutionException)
        {
            if (options.RollbackOnVerificationFailure)
                rollbackPerformed = await RollbackAsync(completed, stepResults, warnings, CancellationToken.None);
            var failed = BuildRun(rollbackPerformed ? RepairRunStatus.RolledBack : RepairRunStatus.Failed);
            return await PersistAsync(failed, progress, CancellationToken.None);
        }
        catch (Exception ex)
        {
            warnings.Add($"Непредвиденная ошибка цели: {ex.Message}");
            if (options.RollbackOnVerificationFailure)
                rollbackPerformed = await RollbackAsync(completed, stepResults, warnings, CancellationToken.None);
            var failed = BuildRun(rollbackPerformed ? RepairRunStatus.RolledBack : RepairRunStatus.Failed);
            return await PersistAsync(failed, progress, CancellationToken.None);
        }

        RepairRun BuildRun(RepairRunStatus status) => new()
        {
            Id = plan.Id,
            Target = plan.Target,
            StartedAt = startedAt,
            CompletedAt = DateTimeOffset.UtcNow,
            Status = status,
            PendingReboot = plan.RequiresReboot,
            Findings = plan.Findings,
            Steps = stepResults.ToArray(),
            Warnings = warnings.ToArray()
        };

        void Report(RepairRunStatus status) => progress?.Report(BuildRun(status));
    }

    private async Task<bool> VerifyFindingsResolvedAsync(RepairPlan plan, CancellationToken cancellationToken)
    {
        if (plan.Findings.Count == 0 || plan.Steps.Count == 0)
            return true;

        var verificationInventory = await inventoryService.CaptureAsync(plan.Target, cancellationToken);
        foreach (var original in plan.Findings.Where(finding =>
                     finding.VerifyResolution && finding.RecommendedActionIds.Count > 0))
        {
            if (!_rules.TryGetValue(original.RuleId, out var rule))
                continue;

            var remaining = await rule.EvaluateAsync(verificationInventory, cancellationToken);
            if (remaining.Any(finding => string.Equals(finding.PrinterName, original.PrinterName, StringComparison.OrdinalIgnoreCase)))
                return false;
        }
        return true;
    }

    private static async Task<bool> RollbackAsync(
        Stack<(IRepairAction Action, RepairActionContext Context, RepairActionCheckpoint Checkpoint, int ResultIndex)> completed,
        List<RepairStepResult> results,
        List<string> warnings,
        CancellationToken cancellationToken)
    {
        var attempted = completed.Count > 0;
        var allSucceeded = attempted;
        while (completed.TryPop(out var item))
        {
            try
            {
                var rollback = await item.Action.RollbackAsync(item.Context, item.Checkpoint, cancellationToken);
                results[item.ResultIndex] = results[item.ResultIndex] with
                {
                    RolledBack = rollback.Success,
                    Output = results[item.ResultIndex].Output.Concat(rollback.Output).ToArray()
                };
                if (!rollback.Success)
                {
                    allSucceeded = false;
                    warnings.Add($"Не удалось откатить '{item.Context.Step.Title}': {rollback.Summary}");
                }
            }
            catch (Exception ex)
            {
                allSucceeded = false;
                warnings.Add($"Ошибка отката '{item.Context.Step.Title}': {ex.Message}");
            }
        }
        return allSucceeded;
    }

    private async Task<RepairRun> PersistAsync(
        RepairRun run,
        IProgress<RepairRun>? progress,
        CancellationToken cancellationToken)
    {
        RepairRun persisted;
        try
        {
            var directory = await reportService.WriteAsync(run, cancellationToken);
            persisted = run with { ReportDirectory = directory };
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            persisted = run with
            {
                Warnings = run.Warnings.Concat([$"Не удалось сохранить отчёт: {ex.Message}"]).ToArray()
            };
        }
        progress?.Report(persisted);
        return persisted;
    }

    private static PrinterInfo? CreatePrinterInfo(PrinterInventory inventory, string? printerName)
    {
        var snapshot = inventory.Printers.FirstOrDefault(printer =>
            string.Equals(printer.Name, printerName, StringComparison.OrdinalIgnoreCase));
        return snapshot is null ? null : new PrinterInfo
        {
            Name = snapshot.Name,
            DriverName = snapshot.DriverName,
            PortName = snapshot.PortName,
            ServerName = inventory.Target.ConnectionName,
            Status = snapshot.Status,
            IsDefault = snapshot.IsDefault,
            IsShared = snapshot.IsShared,
            IsNetwork = snapshot.PortKind is "TCP/IP" or "WSD" or "IPP" or "UNC",
            JobCount = snapshot.JobCount,
            ErrorCodes = snapshot.ErrorCodes
        };
    }

    private static RepairStepResult Failed(RepairStep step, string summary) => new()
    {
        StepId = step.Id,
        ActionId = step.ActionId,
        Summary = summary
    };

    private sealed class RepairExecutionException : Exception;
}
