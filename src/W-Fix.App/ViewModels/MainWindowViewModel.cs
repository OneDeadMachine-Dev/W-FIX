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
    private bool _isFixRunning;
    [ObservableProperty] private string _statusText = "Готов";
    [ObservableProperty] private string _adStatusText = "";
    [ObservableProperty] private string _computerSearchText = "";

    // ── Progress ──────────────────────────────────────────────────────────────
    private CancellationTokenSource? _fixCts;

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
        IsLoading = true;
        StatusText = "Загрузка принтеров...";
        Printers.Clear();
        PrintJobs.Clear();

        try
        {
            await Task.Run(() =>
            {
                var remoteName = SelectedRemoteMachine?.NetBiosName;
                var printers = _wmi.GetPrinters(remoteName);
                Application.Current.Dispatcher.Invoke(() =>
                {
                    foreach (var p in printers) Printers.Add(p);
                });

                // Если AD доступен — добавить опубликованные принтеры (которых нет локально)
                if (_ad.IsDomainAvailable && remoteName == null)
                {
                    var adPrinters = _ad.GetPublishedPrinters();
                    var existingNames = printers.Select(p => p.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        foreach (var ap in adPrinters.Where(p => !existingNames.Contains(p.Name)))
                            Printers.Add(ap);
                    });
                }
            });
        }
        catch (Exception ex)
        {
            AddLog(new LogEntryViewModel(LogLevel.Error, $"Ошибка получения принтеров: {ex.Message}"));
            Serilog.Log.Error(ex, "RefreshPrintersAsync ошибка");
        }

        StatusText = $"Найдено принтеров: {Printers.Count}";
        IsLoading = false;
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
        PrintJobs.Clear();
        if (value == null) return;

        Task.Run(() =>
        {
            var jobs = _wmi.GetPrintJobs(value.Name, SelectedRemoteMachine?.NetBiosName);
            Application.Current.Dispatcher.Invoke(() =>
            {
                foreach (var j in jobs) PrintJobs.Add(j);
            });
        });
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

        IsFixRunning = true;
        LiveLog.Clear();
        _fixCts = new CancellationTokenSource();
        StatusText = $"Выполняется: {SelectedFixer.Name}...";

        AddLog(new LogEntryViewModel(LogLevel.Info, $"▶ Запуск: {SelectedFixer.Name}"));
        if (SelectedRemoteMachine != null)
            AddLog(new LogEntryViewModel(LogLevel.Info, $"  → Удалённая машина: {SelectedRemoteMachine.DisplayName}"));

        var progress = new Progress<LogEntry>(e =>
            Application.Current.Dispatcher.Invoke(() => AddLog(new LogEntryViewModel(e.Level, e.Message))));

        try
        {
            var result = await SelectedFixer.ApplyAsync(
                SelectedPrinter,
                SelectedRemoteMachine?.NetBiosName,
                progress,
                _fixCts.Token);

            var summaryLevel = result.Status == FixStatus.Success ? LogLevel.Success
                             : result.Status == FixStatus.Warning ? LogLevel.Warning
                             : LogLevel.Error;

            AddLog(new LogEntryViewModel(summaryLevel, $"■ Итог: {result.Summary}"));
            StatusText = result.Status == FixStatus.Success ? $"✓ {SelectedFixer.Name} — успешно"
                       : result.Status == FixStatus.Warning ? $"⚠ {SelectedFixer.Name} — предупреждения"
                       : $"✗ {SelectedFixer.Name} — ошибка";

            Serilog.Log.Information("Fixer '{Name}' → {Status}: {Summary}",
                SelectedFixer.Name, result.Status, result.Summary);
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
            Serilog.Log.Error(ex, "Fixer '{Name}' вызвал исключение", SelectedFixer.Name);
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

        IsLoading = true;
        RemoteMachines.Clear();
        StatusText = "Поиск компьютеров в AD...";

        await Task.Run(() =>
        {
            var machines = _ad.GetDomainComputers(ComputerSearchText);
            Application.Current.Dispatcher.Invoke(() =>
            {
                foreach (var m in machines) RemoteMachines.Add(m);
            });
        });

        StatusText = $"Компьютеров в AD: {RemoteMachines.Count}";
        IsLoading = false;
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
