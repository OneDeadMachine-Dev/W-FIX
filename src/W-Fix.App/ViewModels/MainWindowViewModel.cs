using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WFix.Core.Fixers;
using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.App.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    // ── Services ─────────────────────────────────────────────────────────────
    private readonly WmiService _wmi = new();
    private readonly ActiveDirectoryService _ad = new();
    private readonly FixerRegistry _registry = new();
    private readonly SystemStateBackupService _stateBackup = new();

    // ── Collections ───────────────────────────────────────────────────────────
    public ObservableCollection<PrinterInfo> Printers { get; } = [];
    public ObservableCollection<PrintJobInfo> PrintJobs { get; } = [];
    public ObservableCollection<IFixer> AvailableFixers { get; } = [];
    public ObservableCollection<LogEntryViewModel> LiveLog { get; } = [];
    public ObservableCollection<RemoteMachine> RemoteMachines { get; } = [];

    // ── State ─────────────────────────────────────────────────────────────────
    [ObservableProperty] private PrinterInfo? _selectedPrinter;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ApplyFixCommand))]
    private IFixer? _selectedFixer;

    [ObservableProperty] private RemoteMachine? _selectedRemoteMachine;
    [ObservableProperty] private bool _isLoading;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ApplyFixCommand))]
    [NotifyCanExecuteChangedFor(nameof(RestoreLastBackupCommand))]
    private bool _isFixRunning;
    [ObservableProperty] private string _statusText = "Готов";
    [ObservableProperty] private string _adStatusText = "";
    [ObservableProperty] private string _computerSearchText = "";

    // ── Progress ──────────────────────────────────────────────────────────────
    private CancellationTokenSource? _fixCts;
    private CancellationTokenSource? _refreshPrintersCts;
    private CancellationTokenSource? _printJobsCts;
    private CancellationTokenSource? _computerSearchCts;
    private int _loadingOperations;
    private SystemStateBackupResult? _lastBackup;

    public MainWindowViewModel()
    {
        AdStatusText = _ad.IsDomainAvailable
            ? $"AD: {_ad.DomainName}"
            : "AD: домен недоступен (локальный режим)";

        LoadFixers();
        _ = RefreshPrintersAsync();
    }

    // ── Printer Management ────────────────────────────────────────────────────

    [RelayCommand]
    public async Task RefreshPrintersAsync()
    {
        var refreshCts = new CancellationTokenSource();
        Interlocked.Exchange(ref _refreshPrintersCts, refreshCts)?.Cancel();
        var remoteName = SelectedRemoteMachine?.NetBiosName;

        BeginLoading();
        StatusText = "Загрузка принтеров...";
        Printers.Clear();
        PrintJobs.Clear();

        try
        {
            var result = await Task.Run(() =>
            {
                var printers = _wmi.GetPrinters(remoteName);
                IReadOnlyList<PrinterInfo> adPrinters = [];

                // Если AD доступен — добавить опубликованные принтеры (которых нет локально)
                if (_ad.IsDomainAvailable && remoteName == null)
                    adPrinters = _ad.GetPublishedPrinters();

                return (Printers: printers, AdPrinters: adPrinters);
            }, refreshCts.Token);

            refreshCts.Token.ThrowIfCancellationRequested();
            foreach (var printer in result.Printers)
                Printers.Add(printer);

            var existingNames = result.Printers
                .Select(printer => printer.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var printer in result.AdPrinters.Where(printer => !existingNames.Contains(printer.Name)))
                Printers.Add(printer);

            StatusText = $"Найдено принтеров: {Printers.Count}";
        }
        catch (OperationCanceledException) when (refreshCts.IsCancellationRequested)
        {
            // Более новый запрос уже загружает актуальный компьютер.
        }
        catch (Exception ex)
        {
            AddLog(new LogEntryViewModel(LogLevel.Error, $"Ошибка получения принтеров: {ex.Message}"));
            Serilog.Log.Error(ex, "RefreshPrintersAsync ошибка");
            StatusText = "Ошибка загрузки принтеров";
        }
        finally
        {
            Interlocked.CompareExchange(ref _refreshPrintersCts, null, refreshCts);
            EndLoading();
            refreshCts.Dispose();
        }
    }

    [RelayCommand]
    public void OpenLogsFolder()
    {
        var logDir = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "W-Fix", "Logs");
            
        if (System.IO.Directory.Exists(logDir))
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = logDir,
                UseShellExecute = true
            });
        }
    }

    partial void OnSelectedPrinterChanged(PrinterInfo? value)
    {
        var printJobsCts = new CancellationTokenSource();
        Interlocked.Exchange(ref _printJobsCts, printJobsCts)?.Cancel();
        PrintJobs.Clear();
        if (value == null)
        {
            Interlocked.CompareExchange(ref _printJobsCts, null, printJobsCts);
            printJobsCts.Dispose();
            return;
        }

        var remoteName = SelectedRemoteMachine?.NetBiosName;
        _ = LoadPrintJobsAsync(value.Name, remoteName, printJobsCts);
    }

    private async Task LoadPrintJobsAsync(
        string printerName,
        string? remoteMachine,
        CancellationTokenSource printJobsCts)
    {
        try
        {
            var jobs = await Task.Run(
                () => _wmi.GetPrintJobs(printerName, remoteMachine),
                printJobsCts.Token);
            printJobsCts.Token.ThrowIfCancellationRequested();

            foreach (var job in jobs)
                PrintJobs.Add(job);
        }
        catch (OperationCanceledException) when (printJobsCts.IsCancellationRequested)
        {
            // Выбран другой принтер — старый результат больше не нужен.
        }
        catch (Exception ex)
        {
            AddLog(new LogEntryViewModel(LogLevel.Error, $"Ошибка получения очереди печати: {ex.Message}"));
            Serilog.Log.Error(ex, "Не удалось получить очередь принтера '{PrinterName}'", printerName);
        }
        finally
        {
            Interlocked.CompareExchange(ref _printJobsCts, null, printJobsCts);
            printJobsCts.Dispose();
        }
    }

    // ── Fix Execution ─────────────────────────────────────────────────────────

    private void LoadFixers()
    {
        AvailableFixers.Clear();
        foreach (var f in _registry.GetAll())
            AvailableFixers.Add(f);
    }

    [RelayCommand(CanExecute = nameof(CanApplyFix))]
    public async Task ApplyFixAsync()
    {
        if (SelectedFixer == null) return;

        // Если фиксер требует ввода — показываем диалог
        if (SelectedFixer is IInteractiveFixer interactive)
        {
            if (!ShowDriverDialog(SelectedFixer as DriverFixer))
                return; // Пользователь отменил
        }

        var fixer = SelectedFixer;
        var printer = SelectedPrinter;
        var remoteMachine = SelectedRemoteMachine;

        IsFixRunning = true;
        LiveLog.Clear();
        _fixCts = new CancellationTokenSource();
        StatusText = $"Выполняется: {fixer.Name}...";

        AddLog(new LogEntryViewModel(LogLevel.Info, $"▶ Запуск: {fixer.Name}"));
        if (remoteMachine != null)
            AddLog(new LogEntryViewModel(LogLevel.Info, $"  → Удалённая машина: {remoteMachine.DisplayName}"));

        var progress = new Progress<LogEntry>(e =>
            Application.Current.Dispatcher.Invoke(() => AddLog(new LogEntryViewModel(e.Level, e.Message))));

        try
        {
            var backup = await _stateBackup.CreateAsync(
                fixer,
                printer,
                remoteMachine?.NetBiosName,
                _fixCts.Token);

            AddPowerShellOutput(backup.Output);
            if (!backup.Skipped && backup.Success)
            {
                _lastBackup = backup;
                RestoreLastBackupCommand.NotifyCanExecuteChanged();
                AddLog(new LogEntryViewModel(LogLevel.Success,
                    $"↩ Снимок для отката: {backup.BackupDirectory}"));
            }
            else if (!backup.Skipped)
            {
                AddLog(new LogEntryViewModel(LogLevel.Warning,
                    $"⚠ Снимок создать не удалось: {backup.Error}. Исправление будет продолжено."));
            }

            var result = await fixer.ApplyAsync(
                printer,
                remoteMachine?.NetBiosName,
                progress,
                _fixCts.Token);

            var summaryLevel = result.Status == FixStatus.Success ? LogLevel.Success
                             : result.Status == FixStatus.Warning ? LogLevel.Warning
                             : LogLevel.Error;

            AddLog(new LogEntryViewModel(summaryLevel, $"■ Итог: {result.Summary}"));
            StatusText = result.Status == FixStatus.Success ? $"✓ {fixer.Name} — успешно"
                       : result.Status == FixStatus.Warning ? $"⚠ {fixer.Name} — предупреждения"
                       : $"✗ {fixer.Name} — ошибка";

            Serilog.Log.Information("Fixer '{Name}' → {Status}: {Summary}",
                fixer.Name, result.Status, result.Summary);
        }
        catch (OperationCanceledException)
        {
            AddLog(new LogEntryViewModel(LogLevel.Warning, "Операция отменена пользователем."));
            StatusText = "Отменено";
        }
        catch (Exception ex)
        {
            AddLog(new LogEntryViewModel(LogLevel.Error, $"Исключение: {ex.Message}"));
            StatusText = "Ошибка";
            Serilog.Log.Error(ex, "Fixer '{Name}' вызвал исключение", fixer.Name);
        }
        finally
        {
            IsFixRunning = false;
            _fixCts?.Dispose();
            _fixCts = null;
            await RefreshPrintersAsync();
        }
    }

    private bool CanApplyFix() => SelectedFixer != null && !IsFixRunning;

    [RelayCommand(CanExecute = nameof(CanRestoreLastBackup))]
    public async Task RestoreLastBackupAsync()
    {
        var backup = _lastBackup;
        if (backup is not { Success: true, Skipped: false }) return;

        var target = string.IsNullOrWhiteSpace(backup.RemoteMachine)
            ? "локальном компьютере"
            : $"компьютере {backup.RemoteMachine}";
        var answer = MessageBox.Show(
            $"Восстановить состояние реестра на {target} из последнего снимка?\n\n{backup.BackupDirectory}",
            "W-Fix — откат исправления",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);
        if (answer != MessageBoxResult.Yes) return;

        IsFixRunning = true;
        _fixCts = new CancellationTokenSource();
        StatusText = "Восстановление состояния...";
        AddLog(new LogEntryViewModel(LogLevel.Info, $"↩ Запуск отката: {backup.BackupDirectory}"));

        try
        {
            var result = await _stateBackup.RestoreAsync(backup, _fixCts.Token);
            AddPowerShellOutput(result.Output);
            if (result.Success)
            {
                AddLog(new LogEntryViewModel(LogLevel.Success, "■ Откат успешно завершён."));
                StatusText = "✓ Состояние восстановлено";
                _lastBackup = null;
                RestoreLastBackupCommand.NotifyCanExecuteChanged();
            }
            else
            {
                AddLog(new LogEntryViewModel(LogLevel.Error, $"■ Откат не завершён: {result.Error}"));
                StatusText = "✗ Ошибка отката";
            }
        }
        catch (OperationCanceledException)
        {
            AddLog(new LogEntryViewModel(LogLevel.Warning, "Откат отменён пользователем."));
            StatusText = "Отменено";
        }
        catch (Exception ex)
        {
            AddLog(new LogEntryViewModel(LogLevel.Error, $"Ошибка отката: {ex.Message}"));
            StatusText = "✗ Ошибка отката";
            Serilog.Log.Error(ex, "Не удалось восстановить снимок '{BackupDirectory}'", backup.BackupDirectory);
        }
        finally
        {
            IsFixRunning = false;
            _fixCts?.Dispose();
            _fixCts = null;
            await RefreshPrintersAsync();
        }
    }

    private bool CanRestoreLastBackup() =>
        _lastBackup is { Success: true, Skipped: false } && !IsFixRunning;

    [RelayCommand]
    public void CancelFix()
    {
        _fixCts?.Cancel();
        StatusText = "Отмена...";
    }

    /// <summary>
    /// Показывает диалог DriverInstallDialog перед запуском DriverFixer.
    /// Возвращает true если пользователь нажал OK, false — отмена.
    /// </summary>
    private bool ShowDriverDialog(DriverFixer? driverFixer)
    {
        if (driverFixer == null) return false;

        var dialog = new Views.DriverInstallDialog
        {
            Owner = Application.Current.MainWindow
        };

        if (dialog.ShowDialog() != true)
            return false;

        driverFixer.Mode = dialog.SelectedMode;
        switch (dialog.SelectedMode)
        {
            case DriverFixMode.Inf:
                driverFixer.InfPath = dialog.SelectedPath;
                break;
            case DriverFixMode.Unc:
                driverFixer.UncPath = dialog.SelectedPath;
                break;
            case DriverFixMode.Auto:
                // Авто не требует пути
                break;
        }
        return true;
    }

    // ── AD Computer Search ────────────────────────────────────────────────────

    [RelayCommand]
    public async Task SearchComputersAsync()
    {
        if (!_ad.IsDomainAvailable)
        {
            AddLog(new LogEntryViewModel(LogLevel.Warning, "Active Directory недоступен. Работа в локальном режиме."));
            return;
        }

        var searchCts = new CancellationTokenSource();
        Interlocked.Exchange(ref _computerSearchCts, searchCts)?.Cancel();
        var searchText = ComputerSearchText;

        BeginLoading();
        RemoteMachines.Clear();
        StatusText = "Поиск компьютеров в AD...";

        try
        {
            var machines = await Task.Run(
                () => _ad.GetDomainComputers(searchText),
                searchCts.Token);
            searchCts.Token.ThrowIfCancellationRequested();

            foreach (var machine in machines)
                RemoteMachines.Add(machine);
            StatusText = $"Компьютеров в AD: {RemoteMachines.Count}";
        }
        catch (OperationCanceledException) when (searchCts.IsCancellationRequested)
        {
            // Более новый поиск заменил этот запрос.
        }
        catch (Exception ex)
        {
            AddLog(new LogEntryViewModel(LogLevel.Error, $"Ошибка поиска компьютеров: {ex.Message}"));
            StatusText = "Ошибка поиска компьютеров в AD";
            Serilog.Log.Error(ex, "SearchComputersAsync ошибка");
        }
        finally
        {
            Interlocked.CompareExchange(ref _computerSearchCts, null, searchCts);
            EndLoading();
            searchCts.Dispose();
        }
    }

    [RelayCommand]
    public async Task PingRemoteMachineAsync()
    {
        if (SelectedRemoteMachine == null) return;
        using var ping = new System.Net.NetworkInformation.Ping();
        try
        {
            var reply = await ping.SendPingAsync(SelectedRemoteMachine.NetBiosName, 2000);
            SelectedRemoteMachine.IsReachable = reply.Status == System.Net.NetworkInformation.IPStatus.Success;
            StatusText = SelectedRemoteMachine.IsReachable
                ? $"Ping {SelectedRemoteMachine.NetBiosName}: {reply.RoundtripTime} мс"
                : $"Ping {SelectedRemoteMachine.NetBiosName}: недоступен";
        }
        catch (Exception ex)
        {
            StatusText = $"Ping ошибка: {ex.Message}";
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private void AddLog(LogEntryViewModel e)
    {
        LiveLog.Add(e);
        if (LiveLog.Count > 500) LiveLog.RemoveAt(0);
    }

    private void BeginLoading()
    {
        Interlocked.Increment(ref _loadingOperations);
        IsLoading = true;
    }

    private void EndLoading()
    {
        if (Interlocked.Decrement(ref _loadingOperations) == 0)
            IsLoading = false;
    }

    private void AddPowerShellOutput(IEnumerable<string> lines)
    {
        foreach (var line in lines)
        {
            var level = line.StartsWith("[OK]", StringComparison.OrdinalIgnoreCase) ? LogLevel.Success
                : line.StartsWith("[WARN]", StringComparison.OrdinalIgnoreCase) ? LogLevel.Warning
                : line.StartsWith("[ERROR]", StringComparison.OrdinalIgnoreCase) ||
                  line.StartsWith("[EXCEPTION]", StringComparison.OrdinalIgnoreCase) ? LogLevel.Error
                : LogLevel.Info;
            AddLog(new LogEntryViewModel(level, line));
        }
    }
}

/// <summary>ViewModel для строки лога в UI.</summary>
public class LogEntryViewModel(LogLevel level, string message)
{
    public LogLevel Level { get; } = level;
    public string Message { get; } = message;
    public DateTime Timestamp { get; } = DateTime.Now;
    public string Icon => Level switch
    {
        LogLevel.Success => "✅",
        LogLevel.Warning => "⚠️",
        LogLevel.Error   => "❌",
        _                => "ℹ️"
    };
    public string TimeStr => Timestamp.ToString("HH:mm:ss");
}
