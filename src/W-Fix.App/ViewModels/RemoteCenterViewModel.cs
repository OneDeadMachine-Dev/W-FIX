using System.Collections.ObjectModel;
using System.IO;
using System.Net;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WFix.App.Services;
using WFix.Core.Abstractions;
using WFix.Core.Diagnostics;
using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.App.ViewModels;

public partial class RemoteCenterViewModel : ObservableObject, IDisposable
{
    private readonly ActiveDirectoryService _activeDirectory;
    private readonly ICredentialStore _credentialStore;
    private readonly IRemotePreflightService _preflight;
    private readonly IPrinterInventoryService _inventory;
    private readonly IDiagnosticService _diagnostics;
    private readonly IRepairPlanner _planner;
    private readonly IRepairExecutor _executor;
    private readonly IRemoteMaintenanceService _maintenance;
    private readonly ISupportBundleService _supportBundles;
    private readonly IUserPromptService _prompts;
    private CancellationTokenSource? _operationCts;
    private CredentialReference? _credential;

    public ObservableCollection<RemoteTargetViewModel> Targets { get; } = [];
    public ObservableCollection<RemoteMachine> ActiveDirectoryResults { get; } = [];
    public ObservableCollection<DiagnosticFinding> Findings { get; } = [];
    public ObservableCollection<RepairStepViewModel> PlannedSteps { get; } = [];
    public ObservableCollection<RepairRun> Runs { get; } = [];

    [ObservableProperty] private string _manualTargets = "";
    [ObservableProperty] private string _activeDirectorySearch = "";
    [ObservableProperty] private string _credentialUserName = "";
    [ObservableProperty] private string _credentialStatus = "Текущая учётная запись Windows (Kerberos)";
    [ObservableProperty] private int _maxConcurrency = 3;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusText = "Добавьте компьютеры и запустите preflight.";

    public RemoteCenterViewModel(
        ActiveDirectoryService activeDirectory,
        ICredentialStore credentialStore,
        IRemotePreflightService preflight,
        IPrinterInventoryService inventory,
        IDiagnosticService diagnostics,
        IRepairPlanner planner,
        IRepairExecutor executor,
        IRemoteMaintenanceService maintenance,
        ISupportBundleService supportBundles,
        IUserPromptService prompts)
    {
        _activeDirectory = activeDirectory;
        _credentialStore = credentialStore;
        _preflight = preflight;
        _inventory = inventory;
        _diagnostics = diagnostics;
        _planner = planner;
        _executor = executor;
        _maintenance = maintenance;
        _supportBundles = supportBundles;
        _prompts = prompts;
    }

    public async Task SaveCredentialAsync(string password)
    {
        if (string.IsNullOrWhiteSpace(CredentialUserName))
            throw new InvalidOperationException("Укажите доменную учётную запись.");
        if (string.IsNullOrEmpty(password))
            throw new InvalidOperationException("Введите пароль.");

        var key = new string(CredentialUserName
            .Where(character => char.IsLetterOrDigit(character) || character is '-' or '_' or '.' or '@')
            .ToArray());
        if (string.IsNullOrWhiteSpace(key))
            key = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(CredentialUserName)))[..16];

        var reference = new CredentialReference($"W-Fix/domain/{key}", CredentialUserName);
        await _credentialStore.SaveAsync(reference, new NetworkCredential(CredentialUserName, password));
        _credential = reference;
        ApplyCredential(reference);
        CredentialStatus = $"Credential Manager: {CredentialUserName}";
        StatusText = "Альтернативная учётная запись сохранена средствами Windows.";
    }

    public Task<string> ExportSupportBundleAsync(string outputPath) =>
        _supportBundles.ExportAsync(Runs.ToArray(), outputPath);

    [RelayCommand]
    private void UseCurrentCredential()
    {
        _credential = null;
        ApplyCredential(null);
        CredentialStatus = "Текущая учётная запись Windows (Kerberos)";
    }

    [RelayCommand]
    private async Task DeleteSavedCredentialAsync()
    {
        if (_credential is null)
            return;
        await _credentialStore.DeleteAsync(_credential);
        UseCurrentCredential();
        StatusText = "Сохранённая учётная запись удалена из Credential Manager.";
    }

    [RelayCommand]
    private void AddManualTargets()
    {
        foreach (var target in TargetParser.ParseManual(ManualTargets, _credential))
            AddTarget(target);
        ManualTargets = "";
        StatusText = $"Целей: {Targets.Count}.";
    }

    [RelayCommand]
    private async Task SearchActiveDirectoryAsync()
    {
        if (!_activeDirectory.IsDomainAvailable)
        {
            StatusText = "Active Directory недоступен.";
            return;
        }

        IsBusy = true;
        try
        {
            var result = await Task.Run(() => _activeDirectory.GetDomainComputers(ActiveDirectorySearch));
            ActiveDirectoryResults.Clear();
            foreach (var machine in result.Take(500))
                ActiveDirectoryResults.Add(machine);
            StatusText = $"Найдено в AD: {ActiveDirectoryResults.Count}.";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void AddActiveDirectoryTarget(RemoteMachine? machine)
    {
        if (machine is not null)
            AddTarget(TargetParser.FromActiveDirectory(machine, _credential));
    }

    [RelayCommand]
    private void RemoveTarget(RemoteTargetViewModel? target)
    {
        if (target is not null)
            Targets.Remove(target);
    }

    [RelayCommand]
    private async Task RunPreflightAsync()
    {
        if (Targets.Count == 0 || IsBusy)
            return;

        await RunOperationAsync(async cancellationToken =>
        {
            StatusText = "Проверка DNS, WinRM, полномочий и служб...";
            using var gate = new SemaphoreSlim(Math.Min(5, Math.Max(1, MaxConcurrency)));
            var checks = Targets.Select(async item =>
            {
                await gate.WaitAsync(cancellationToken);
                try
                {
                    item.State = "Preflight...";
                    var report = await _preflight.CheckAsync(item.Target, cancellationToken);
                    return (item, report);
                }
                finally
                {
                    gate.Release();
                }
            }).ToArray();

            foreach (var (item, report) in await Task.WhenAll(checks))
            {
                item.Capabilities = report;
                item.State = report.CanRepairMachine ? "Готов к ремонту" : report.CanInspect ? "Только диагностика" : "Недоступен";
                item.Details = BuildCapabilitySummary(report);
            }
            StatusText = $"Preflight завершён: {Targets.Count(item => item.Capabilities?.CanInspect == true)} из {Targets.Count} доступны.";
        });
    }

    [RelayCommand]
    private async Task DiagnoseAsync()
    {
        if (Targets.Count == 0 || IsBusy)
            return;

        if (Targets.Any(target => target.Capabilities is null))
            await RunPreflightAsync();

        await RunOperationAsync(async cancellationToken =>
        {
            Findings.Clear();
            PlannedSteps.Clear();
            StatusText = "Инвентаризация и диагностические правила...";
            using var gate = new SemaphoreSlim(Math.Min(3, Math.Max(1, MaxConcurrency)));
            var tasks = Targets.Where(item => item.Capabilities?.CanInspect == true).Select(async item =>
            {
                await gate.WaitAsync(cancellationToken);
                try
                {
                    item.State = "Диагностика...";
                    var inventory = await _inventory.CaptureAsync(item.Target, cancellationToken);
                    var findings = await _diagnostics.DiagnoseAsync(inventory, cancellationToken);
                    var plan = _planner.CreatePlan(item.Target, findings, item.Capabilities!);
                    return (item, inventory, findings, plan, error: (string?)null);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    return (item, inventory: (PrinterInventory?)null, findings: (IReadOnlyList<DiagnosticFinding>)[], plan: (RepairPlan?)null, error: ex.Message);
                }
                finally
                {
                    gate.Release();
                }
            }).ToArray();

            foreach (var result in await Task.WhenAll(tasks))
            {
                result.item.Inventory = result.inventory;
                result.item.Plan = result.plan;
                result.item.State = result.error is null ? $"Найдено: {result.findings.Count}" : "Ошибка диагностики";
                result.item.Details = result.error ?? result.item.Details;
                foreach (var finding in result.findings)
                    Findings.Add(finding);
                if (result.plan is not null)
                    foreach (var step in result.plan.Steps)
                        PlannedSteps.Add(new RepairStepViewModel(result.plan.Target, step));
            }
            StatusText = $"Диагностика завершена: {Findings.Count} находок, {PlannedSteps.Count} действий.";
        });
    }

    [RelayCommand]
    private async Task ExecutePlansAsync()
    {
        if (IsBusy)
            return;
        var plans = Targets.Select(target => target.Plan).Where(plan => plan is { Steps.Count: > 0 }).Cast<RepairPlan>().ToArray();
        if (plans.Length == 0)
        {
            StatusText = "Нет готовых планов ремонта. Сначала запустите диагностику.";
            return;
        }

        var actionCount = plans.Sum(plan => plan.Steps.Count);
        if (!_prompts.Confirm(
                $"Будет выполнено действий: {actionCount}\nКомпьютеров: {plans.Length}\n\nПеред изменениями W-Fix создаст снимки и проверит результат.",
                "Запустить пакетный ремонт?"))
            return;

        if (plans.Any(plan => plan.RequiresAdditionalConfirmation) && !_prompts.Confirm(
                "План содержит необратимые действия (например, работу с пакетами драйверов). Проверьте список ещё раз. Продолжить?",
                "Дополнительный барьер безопасности",
                danger: true))
            return;

        await RunOperationAsync(async cancellationToken =>
        {
            Runs.Clear();
            StatusText = "Выполняется пакетный ремонт...";
            var progress = new Progress<RepairRun>(run =>
            {
                var target = Targets.FirstOrDefault(item => item.Target.Id == run.Target.Id);
                if (target is not null)
                {
                    target.Run = run;
                    target.State = run.Status.ToString();
                }
            });
            var runs = await _executor.ExecuteBatchAsync(
                plans,
                new RepairBatchOptions { MaxConcurrency = Math.Clamp(MaxConcurrency, 1, 10) },
                progress,
                cancellationToken);
            foreach (var run in runs)
                Runs.Add(run);
            StatusText = $"Пакет завершён: успешно {runs.Count(run => run.Status == RepairRunStatus.Succeeded)}, с ошибкой/откатом {runs.Count(run => run.Status is RepairRunStatus.Failed or RepairRunStatus.RolledBack)}.";
        });
    }

    [RelayCommand]
    private async Task RestartPendingAsync()
    {
        var pending = Runs.Where(run => run.PendingReboot && run.Status == RepairRunStatus.Succeeded).ToArray();
        if (pending.Length == 0)
        {
            StatusText = "Нет успешно отремонтированных машин, ожидающих перезагрузку.";
            return;
        }
        if (!_prompts.Confirm($"Отправить команду перезагрузки на {pending.Length} компьютеров?", "Перезагрузка", danger: true))
            return;

        await RunOperationAsync(async cancellationToken =>
        {
            foreach (var run in pending)
            {
                var result = await _maintenance.RestartAsync(run.Target, cancellationToken);
                var target = Targets.FirstOrDefault(item => item.Target.Id == run.Target.Id);
                if (target is not null)
                    target.State = result.Success ? "Перезагрузка отправлена" : $"Ошибка перезагрузки: {result.Error}";
            }
        });
    }

    [RelayCommand]
    private void Cancel() => _operationCts?.Cancel();

    [RelayCommand]
    private void OpenReport(RepairRun? run)
    {
        var reportDirectory = run?.ReportDirectory;
        if (reportDirectory is null || !Directory.Exists(reportDirectory))
            return;
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = reportDirectory,
            UseShellExecute = true
        });
    }

    private async Task RunOperationAsync(Func<CancellationToken, Task> operation)
    {
        if (IsBusy)
            return;
        _operationCts?.Dispose();
        _operationCts = new CancellationTokenSource();
        IsBusy = true;
        try
        {
            await operation(_operationCts.Token);
        }
        catch (OperationCanceledException)
        {
            StatusText = "Операция отменена. Выполненные обратимые шаги откатываются.";
        }
        catch (Exception ex)
        {
            StatusText = $"Ошибка: {ex.Message}";
            Serilog.Log.Error(ex, "Remote Center operation failed");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void AddTarget(TargetDescriptor target)
    {
        if (Targets.Any(item => item.Target.Id.Equals(target.Id, StringComparison.OrdinalIgnoreCase)))
            return;
        Targets.Add(new RemoteTargetViewModel(target));
    }

    private void ApplyCredential(CredentialReference? credential)
    {
        foreach (var item in Targets)
        {
            item.Target = item.Target with { Credential = credential };
            item.Capabilities = null;
            item.Plan = null;
            item.State = "Требуется preflight";
        }
    }

    private static string BuildCapabilitySummary(RemoteCapabilityReport report) =>
        $"DNS {(report.DnsResolved ? "✓" : "✗")} · Ping {(report.PingResponded ? "✓" : "—")} · WinRM {(report.WinRmAvailable ? "✓" : "✗")} · Admin {(report.IsAdministrator ? "✓" : "✗")} · Spooler {(report.SpoolerAvailable ? "✓" : "✗")}";

    public void Dispose()
    {
        _operationCts?.Cancel();
        _operationCts?.Dispose();
    }
}

public partial class RemoteTargetViewModel(TargetDescriptor target) : ObservableObject
{
    [ObservableProperty] private TargetDescriptor _target = target;
    [ObservableProperty] private string _state = "Ожидает";
    [ObservableProperty] private string _details = "";
    [ObservableProperty] private RemoteCapabilityReport? _capabilities;
    [ObservableProperty] private PrinterInventory? _inventory;
    [ObservableProperty] private RepairPlan? _plan;
    [ObservableProperty] private RepairRun? _run;
    public string DisplayName => Target.ConnectionName;

    partial void OnTargetChanged(TargetDescriptor value) => OnPropertyChanged(nameof(DisplayName));
}

public sealed record RepairStepViewModel(TargetDescriptor Target, RepairStep Step)
{
    public string Computer => Target.ConnectionName;
    public string Title => Step.Title;
    public string Printer => Step.PrinterName ?? "Вся система";
    public string Risk => Step.Risk.ToString();
}
