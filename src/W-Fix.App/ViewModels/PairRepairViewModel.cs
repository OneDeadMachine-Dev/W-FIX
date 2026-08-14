using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WFix.App.Services;
using WFix.Core.Abstractions;
using WFix.Core.Models;
using WFix.Core.Pairing;
using WFix.Core.Services;

namespace WFix.App.ViewModels;

public partial class PairRepairViewModel : ObservableObject, IAsyncDisposable
{
    private readonly IPairSessionTransport _transport;
    private readonly IPairFirewallLeaseService _firewall;
    private readonly IPairFileService _files;
    private readonly IPairInventoryService _inventory;
    private readonly IPairDiagnosticService _diagnostics;
    private readonly IPairRepairPlanner _planner;
    private readonly IPairActionDispatcher _localDispatcher;
    private readonly IPairAgentCommandLoop _agentLoop;
    private readonly IPairRunReportService _reports;
    private readonly INetworkCredentialProvisioner _credentials;
    private readonly IUserPromptService _prompts;
    private readonly InteractiveUserPowerShellService _interactiveUserPowerShell;
    private readonly CancellationTokenSource _lifetime = new();
    private IPairHost? _host;
    private IPairSession? _session;
    private IAsyncDisposable? _firewallLease;
    private Task? _agentTask;
    private PairInvitation? _invitation;
    private PairEndpointSnapshot? _hostSnapshot;
    private PairEndpointSnapshot? _clientSnapshot;
    private PairRepairPlan? _plan;
    private IReadOnlyDictionary<PairEndpointRole, TargetDescriptor>? _executionTargets;
    private bool _disposed;

    public ObservableCollection<PairDiagnosticFinding> Findings { get; } = [];
    public ObservableCollection<PairRepairStep> Steps { get; } = [];
    public ObservableCollection<string> Activity { get; } = [];

    [ObservableProperty] private PairEndpointRole _localRole = PairEndpointRole.Host;
    [ObservableProperty] private PairTransportMode _transportMode = PairTransportMode.LiveLan;
    [ObservableProperty] private string _peerComputerName = "";
    [ObservableProperty] private string _domainHostComputerName = "";
    [ObservableProperty] private string _printerName = "";
    [ObservableProperty] private string _shareName = "";
    [ObservableProperty] private string _confirmationCode = "------";
    [ObservableProperty] private string _statusText = "Выберите роль этого ПК.";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _isConnected;
    [ObservableProperty] private bool _isApproved;
    [ObservableProperty] private bool _includeExpertActions;
    [ObservableProperty] private bool _physicalTestConfirmed;
    [ObservableProperty] private string _networkCredentialUser = "";
    [ObservableProperty] private PairRun? _lastRun;

    public string LocalComputerName => Environment.MachineName;
    public string RoleText => LocalRole == PairEndpointRole.Host ? "Хост принтера" : "Клиент";
    public bool HasInvitation => _invitation is not null;
    public bool HasPlan => _plan is { Steps.Count: > 0 };
    public bool CanExecute => HasPlan && !IsBusy;

    public PairRepairViewModel(
        IPairSessionTransport transport,
        IPairFirewallLeaseService firewall,
        IPairFileService files,
        IPairInventoryService inventory,
        IPairDiagnosticService diagnostics,
        IPairRepairPlanner planner,
        IPairActionDispatcher localDispatcher,
        IPairAgentCommandLoop agentLoop,
        IPairRunReportService reports,
        INetworkCredentialProvisioner credentials,
        IUserPromptService prompts,
        InteractiveUserPowerShellService interactiveUserPowerShell)
    {
        _transport = transport;
        _firewall = firewall;
        _files = files;
        _inventory = inventory;
        _diagnostics = diagnostics;
        _planner = planner;
        _localDispatcher = localDispatcher;
        _agentLoop = agentLoop;
        _reports = reports;
        _credentials = credentials;
        _prompts = prompts;
        _interactiveUserPowerShell = interactiveUserPowerShell;
        Activity.Add("W-Fix не включает SMB1 и не устанавливает одинаковые пароли пользователей.");
    }

    [RelayCommand]
    private void UseHostRole()
    {
        EnsureNotConnected();
        LocalRole = PairEndpointRole.Host;
        StatusText = "Укажите имя клиентского ПК и локальный принтер.";
        OnPropertyChanged(nameof(RoleText));
    }

    [RelayCommand]
    private void UseClientRole()
    {
        EnsureNotConnected();
        LocalRole = PairEndpointRole.Client;
        StatusText = "Импортируйте приглашение хоста.";
        OnPropertyChanged(nameof(RoleText));
    }

    [RelayCommand]
    private async Task StartHostAsync()
    {
        EnsureRole(PairEndpointRole.Host);
        RequirePeer();
        if (string.IsNullOrWhiteSpace(PrinterName))
            throw new InvalidOperationException("Укажите имя локального принтера на хосте.");
        await RunBusyAsync(async cancellationToken =>
        {
            await DisposeConnectionAsync();
            await _firewall.CleanupStaleAsync(cancellationToken);
            _host = await _transport.StartHostAsync(new PairHostOptions
            {
                HostComputerName = Environment.MachineName,
                PrinterName = PrinterName.Trim(),
                ShareName = NullIfEmpty(ShareName),
                ExpectedClientComputerName = PeerComputerName.Trim(),
                InvitationLifetime = TimeSpan.FromMinutes(15)
            }, cancellationToken);
            try
            {
                _firewallLease = await _firewall.OpenAsync(
                    _host.Invitation.SessionId,
                    _host.Invitation.Port,
                    Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule!.FileName,
                    cancellationToken);
            }
            catch
            {
                await _host.DisposeAsync();
                _host = null;
                throw;
            }
            _invitation = _host.Invitation;
            ConfirmationCode = _invitation.ConfirmationCode;
            TransportMode = PairTransportMode.LiveLan;
            Activity.Add($"Приглашение создано до {_invitation.ExpiresAt.LocalDateTime:t}; порт {_invitation.Port}.");
            StatusText = "Экспортируйте .wfixpair и нажмите «Ждать подключение».";
            OnPropertyChanged(nameof(HasInvitation));
        });
    }

    public async Task ExportInvitationAsync(string path)
    {
        if (_invitation is null) throw new InvalidOperationException("Сначала создайте приглашение.");
        await _files.WriteInvitationAsync(path, _invitation, _lifetime.Token);
        Activity.Add($"Приглашение сохранено: {Path.GetFileName(path)}");
    }

    [RelayCommand]
    private async Task WaitForPartnerAsync()
    {
        EnsureRole(PairEndpointRole.Host);
        if (_host is null) throw new InvalidOperationException("Сначала создайте приглашение.");
        await RunBusyAsync(async cancellationToken =>
        {
            StatusText = "Ожидание клиента…";
            _session = await _host.AcceptAsync(cancellationToken);
            PeerComputerName = _session.PeerComputerName;
            IsConnected = true;
            ConfirmationCode = _session.ConfirmationCode;
            StatusText = "Сверьте код на обоих ПК и подтвердите соединение.";
            Activity.Add("TLS-соединение установлено; изменения пока запрещены.");
        });
    }

    public async Task ImportInvitationAsync(string path)
    {
        EnsureRole(PairEndpointRole.Client);
        await RunBusyAsync(async cancellationToken =>
        {
            await DisposeConnectionAsync();
            _invitation = await _files.ReadInvitationAsync(path, cancellationToken);
            PeerComputerName = _invitation.HostComputerName;
            PrinterName = _invitation.PrinterName ?? "";
            ShareName = _invitation.ShareName ?? "";
            _session = await _transport.JoinAsync(_invitation, cancellationToken);
            IsConnected = true;
            ConfirmationCode = _session.ConfirmationCode;
            TransportMode = PairTransportMode.LiveLan;
            StatusText = "Сверьте код на обоих ПК и подтвердите соединение.";
            Activity.Add("TLS-соединение установлено, pinning временного ECDSA-ключа пройден.");
            OnPropertyChanged(nameof(HasInvitation));
        });
    }

    [RelayCommand]
    private async Task ApproveAsync()
    {
        if (_session is null || !IsConnected) throw new InvalidOperationException("Нет подключённой pairing-сессии.");
        await RunBusyAsync(async cancellationToken =>
        {
            StatusText = "Ожидание подтверждения второго ПК…";
            if (!await _session.ApproveAsync(true, cancellationToken))
                throw new InvalidOperationException("Вторая сторона отклонила pairing-сессию.");
            IsApproved = true;
            Activity.Add("Одинаковый код подтверждён на обоих ПК.");
            if (LocalRole == PairEndpointRole.Host)
            {
                _hostSnapshot = await _inventory.CaptureAsync(
                    TargetDescriptor.Local(), PairEndpointRole.Host, PeerComputerName.Trim(),
                    PrinterName.Trim(), NullIfEmpty(ShareName), cancellationToken);
                await _session.SendAsync(PairMessageKind.Snapshot, _hostSnapshot, cancellationToken);
                StatusText = "Хост готов. На клиенте можно запускать диагностику и ремонт.";
                _agentTask = RunAgentAsync(_session, _lifetime.Token);
            }
            else
            {
                _clientSnapshot = await _inventory.CaptureAsync(
                    TargetDescriptor.Local(), PairEndpointRole.Client, PeerComputerName.Trim(),
                    null, NullIfEmpty(ShareName), cancellationToken);
                _hostSnapshot = await _session.ReceiveAsync<PairEndpointSnapshot>(PairMessageKind.Snapshot, cancellationToken);
                _executionTargets = LocalTargets();
                await BuildPlanAsync(PairTransportMode.LiveLan, cancellationToken);
            }
        });
    }

    public async Task SaveNetworkCredentialAsync(string password)
    {
        EnsureRole(PairEndpointRole.Client);
        var host = _invitation?.HostComputerName ?? PeerComputerName.Trim();
        if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(NetworkCredentialUser) || string.IsNullOrEmpty(password))
            throw new InvalidOperationException("Укажите хост, пользователя HOST\\User и пароль.");
        await _credentials.SaveForHostAsync(host, new NetworkCredential(NetworkCredentialUser.Trim(), password), _lifetime.Token);
        Activity.Add($"Windows network credential сохранён только для {host}.");
        StatusText = "Учётная запись сохранена в Windows Credential Manager; пароль не попал в W-Fix.";
    }

    public async Task ExportOfflineSnapshotAsync(string path)
    {
        RequirePeer();
        await RunBusyAsync(async cancellationToken =>
        {
            var snapshot = await _inventory.CaptureAsync(
                TargetDescriptor.Local(), LocalRole, PeerComputerName.Trim(),
                NullIfEmpty(PrinterName), NullIfEmpty(ShareName), cancellationToken);
            await _files.WriteOfflineSnapshotAsync(path, snapshot, cancellationToken);
            Activity.Add($"Подписанный offline snapshot сохранён: {Path.GetFileName(path)}");
            StatusText = "Передайте файл только второму выбранному ПК.";
        });
    }

    public async Task ImportOfflineSnapshotAsync(string path)
    {
        RequirePeer();
        await RunBusyAsync(async cancellationToken =>
        {
            var peer = await _files.ReadOfflineSnapshotAsync(path, cancellationToken);
            if (peer.Endpoint.Role == LocalRole)
                throw new InvalidDataException("Импортирован снимок той же роли. Нужен снимок второго ПК.");
            var local = await _inventory.CaptureAsync(
                TargetDescriptor.Local(), LocalRole, peer.Endpoint.ComputerName,
                NullIfEmpty(PrinterName), NullIfEmpty(ShareName), cancellationToken);
            _hostSnapshot = LocalRole == PairEndpointRole.Host ? local : peer;
            _clientSnapshot = LocalRole == PairEndpointRole.Client ? local : peer;
            _executionTargets = LocalTargets();
            TransportMode = PairTransportMode.Offline;
            await BuildPlanAsync(PairTransportMode.Offline, cancellationToken);
            StatusText = "Offline-план готов только для этого ПК; общий автоматический rollback невозможен.";
        });
    }

    [RelayCommand]
    private async Task DiagnoseDomainPairAsync()
    {
        var hostName = DomainHostComputerName.Trim();
        var clientName = PeerComputerName.Trim();
        if (string.IsNullOrWhiteSpace(hostName) || string.IsNullOrWhiteSpace(clientName))
            throw new InvalidOperationException("Укажите имена доменных Host и Client.");
        if (string.Equals(hostName, clientName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Host и Client должны быть разными компьютерами.");
        if (string.IsNullOrWhiteSpace(PrinterName))
            throw new InvalidOperationException("Укажите точное имя очереди на Host.");

        await RunBusyAsync(async cancellationToken =>
        {
            var hostTarget = RemoteTarget(hostName);
            var clientTarget = RemoteTarget(clientName);
            var hostTask = _inventory.CaptureAsync(
                hostTarget, PairEndpointRole.Host, clientName, PrinterName.Trim(), NullIfEmpty(ShareName), cancellationToken);
            var clientTask = _inventory.CaptureAsync(
                clientTarget, PairEndpointRole.Client, hostName, null, NullIfEmpty(ShareName), cancellationToken);
            await Task.WhenAll(hostTask, clientTask);
            _hostSnapshot = await hostTask;
            _clientSnapshot = await clientTask;
            _executionTargets = new Dictionary<PairEndpointRole, TargetDescriptor>
            {
                [PairEndpointRole.Host] = hostTarget,
                [PairEndpointRole.Client] = clientTarget
            };
            TransportMode = PairTransportMode.DomainRemote;
            await BuildPlanAsync(PairTransportMode.DomainRemote, cancellationToken);
            Activity.Add($"Доменная инвентаризация: {hostName} ↔ {clientName} через WinRM/Kerberos.");
        });
    }

    [RelayCommand]
    private async Task RebuildPlanAsync()
    {
        if (_hostSnapshot is null || _clientSnapshot is null) throw new InvalidOperationException("Сначала соберите снимки обоих ПК.");
        await RunBusyAsync(cancellationToken => BuildPlanAsync(TransportMode, cancellationToken));
    }

    [RelayCommand]
    private async Task ExecutePlanAsync()
    {
        if (_plan is null || _plan.Steps.Count == 0) throw new InvalidOperationException("План ремонта пуст.");
        if (_plan.RequiresAdditionalConfirmation && !_prompts.Confirm(
                "План содержит необратимое или экспертное действие. Снимки будут созданы до изменений. Продолжить?",
                "Pair Repair — повышенный риск", true))
            return;
        if (!_prompts.Confirm(
                $"Будет выполнено действий: {_plan.Steps.Count}. Результат проверяется повторно, при сбое запускается rollback.",
                "Запустить Pair Repair?"))
            return;

        await RunBusyAsync(async cancellationToken =>
        {
            IPairActionDispatcher dispatcher = _localDispatcher;
            if (TransportMode == PairTransportMode.LiveLan)
            {
                EnsureRole(PairEndpointRole.Client);
                dispatcher = new PairSessionActionDispatcher(
                    _session ?? throw new InvalidOperationException("Live pairing-сессия закрыта."),
                    PairEndpointRole.Host,
                    _localDispatcher);
            }
            var executor = new PairRepairExecutor(dispatcher, _reports);
            var targets = _executionTargets ?? throw new InvalidOperationException("Не определены цели PairRun.");
            LastRun = await executor.ExecuteAsync(_plan, targets, cancellationToken: cancellationToken);
            foreach (var step in LastRun.Steps)
                Activity.Add($"{step.Endpoint}: {step.Summary} (verify={step.Verified}, rollback={step.RolledBack})");
            StatusText = LastRun.Status == PairRunStatus.Succeeded
                ? "Ремонт выполнен и подтверждён повторной диагностикой. Напечатайте тестовую страницу."
                : $"Ремонт завершён со статусом {LastRun.Status}. Откройте отчёт для деталей.";
        });
    }

    [RelayCommand]
    private async Task PrintTestPageAsync()
    {
        var host = _invitation?.HostComputerName ?? _hostSnapshot?.Endpoint.ComputerName ?? PeerComputerName.Trim();
        var share = _plan?.Steps.SelectMany(step => step.Parameters).FirstOrDefault(pair => pair.Key == "shareName").Value
                    ?? _hostSnapshot?.PrinterShareName ?? NullIfEmpty(ShareName);
        if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(share))
            throw new InvalidOperationException("Не удалось определить имя общей очереди.");
        var connection = $"\\\\{host}\\{share}";
        if (TransportMode == PairTransportMode.DomainRemote)
        {
            var client = _executionTargets?[PairEndpointRole.Client].ConnectionName
                         ?? throw new InvalidOperationException("Не определён доменный Client.");
            var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(connection));
            var script = "$n=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encoded + "'));" +
                         "Start-Process (Join-Path $env:SystemRoot 'System32\\rundll32.exe') -ArgumentList @('printui.dll,PrintUIEntry','/k','/n',$n) -Wait -NoNewWindow";
            var result = await _interactiveUserPowerShell.RunRemoteAsync(client, script, _lifetime.Token, TimeSpan.FromMinutes(2));
            if (!result.Success) throw new InvalidOperationException(result.Error ?? "Не удалось отправить тестовую страницу на доменном Client.");
        }
        else
        {
            var start = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "rundll32.exe"))
            {
                UseShellExecute = false,
                CreateNoWindow = true
            };
            start.ArgumentList.Add("printui.dll,PrintUIEntry");
            start.ArgumentList.Add("/k");
            start.ArgumentList.Add("/n");
            start.ArgumentList.Add(connection);
            using var process = Process.Start(start) ?? throw new InvalidOperationException("Не удалось запустить печать тестовой страницы.");
            await process.WaitForExitAsync(_lifetime.Token);
        }
        PhysicalTestConfirmed = _prompts.Confirm("Тестовая страница физически напечаталась?", "Физическая проверка");
        StatusText = PhysicalTestConfirmed
            ? "Готово: очередь подключена, диагностика чистая, физическая печать подтверждена."
            : "Windows отправила тестовую страницу, но физическая печать не подтверждена — требуется проверка устройства.";
        Activity.Add($"Физическая проверка: {(PhysicalTestConfirmed ? "успешно" : "не подтверждена") }.");
    }

    [RelayCommand]
    private void Cancel() => _lifetime.Cancel();

    private async Task BuildPlanAsync(PairTransportMode mode, CancellationToken cancellationToken)
    {
        var host = _hostSnapshot ?? throw new InvalidOperationException("Нет снимка хоста.");
        var client = _clientSnapshot ?? throw new InvalidOperationException("Нет снимка клиента.");
        var findings = await _diagnostics.DiagnoseAsync(host, client, cancellationToken);
        var fullPlan = _planner.CreatePlan(host, client, mode, findings, IncludeExpertActions);
        _plan = mode == PairTransportMode.Offline ? LocalOnly(fullPlan, LocalRole) : fullPlan;
        Findings.Clear();
        foreach (var finding in findings) Findings.Add(finding);
        Steps.Clear();
        foreach (var step in _plan.Steps) Steps.Add(step);
        StatusText = _plan.Steps.Count == 0
            ? "Диагностика не нашла действий для выбранного режима."
            : $"Найдено причин: {findings.Count}; действий в плане: {_plan.Steps.Count}.";
        OnPropertyChanged(nameof(HasPlan));
        OnPropertyChanged(nameof(CanExecute));
    }

    private static PairRepairPlan LocalOnly(PairRepairPlan plan, PairEndpointRole role)
    {
        var filtered = plan.Steps.Where(step => step.Endpoint == role).ToArray();
        var normalized = filtered.Select((step, index) => step with
        {
            DependsOn = index == 0 ? [] : [filtered[index - 1].Id]
        }).ToArray();
        return plan with { Steps = normalized };
    }

    private static IReadOnlyDictionary<PairEndpointRole, TargetDescriptor> LocalTargets() =>
        new Dictionary<PairEndpointRole, TargetDescriptor>
        {
            [PairEndpointRole.Host] = TargetDescriptor.Local(),
            [PairEndpointRole.Client] = TargetDescriptor.Local()
        };

    private static TargetDescriptor RemoteTarget(string computerName) => new()
    {
        Id = computerName.ToUpperInvariant(),
        ComputerName = computerName,
        Source = TargetSource.Manual
    };

    private async Task RunAgentAsync(IPairSession session, CancellationToken cancellationToken)
    {
        try
        {
            await _agentLoop.RunAsync(session, cancellationToken);
            StatusText = "Клиент завершил PairRun; host-side изменения подтверждены.";
            Activity.Add("Host agent получил commit PairRun.");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            Activity.Add("Host agent остановлен; незавершённые обратимые действия откатываются.");
        }
        catch (Exception ex)
        {
            StatusText = "Связь потеряна; host-side rollback выполнен для незавершённого плана.";
            Activity.Add("Host agent: " + ex.Message);
        }
        finally
        {
            if (_firewallLease is not null)
            {
                await _firewallLease.DisposeAsync();
                _firewallLease = null;
            }
        }
    }

    private async Task RunBusyAsync(Func<CancellationToken, Task> operation)
    {
        if (IsBusy) return;
        IsBusy = true;
        try
        {
            await operation(_lifetime.Token);
        }
        catch (OperationCanceledException)
        {
            StatusText = "Операция отменена; незавершённый обратимый план откатывается.";
        }
        catch (Exception ex)
        {
            StatusText = ex.Message;
            Activity.Add("Ошибка: " + ex.Message);
        }
        finally
        {
            IsBusy = false;
            OnPropertyChanged(nameof(CanExecute));
        }
    }

    private async Task DisposeConnectionAsync()
    {
        if (_session is not null) { await _session.DisposeAsync(); _session = null; }
        if (_firewallLease is not null) { await _firewallLease.DisposeAsync(); _firewallLease = null; }
        if (_host is not null) { await _host.DisposeAsync(); _host = null; }
        IsConnected = false;
        IsApproved = false;
    }

    private void EnsureRole(PairEndpointRole role)
    {
        if (LocalRole != role) throw new InvalidOperationException($"Это действие предназначено для роли {role}.");
    }

    private void RequirePeer()
    {
        if (string.IsNullOrWhiteSpace(PeerComputerName)) throw new InvalidOperationException("Укажите имя второго ПК.");
    }

    private void EnsureNotConnected()
    {
        if (IsConnected) throw new InvalidOperationException("Нельзя менять роль во время активной сессии.");
    }

    private static string? NullIfEmpty(string? value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;
        _lifetime.Cancel();
        await DisposeConnectionAsync();
        if (_agentTask is not null)
        {
            try { await _agentTask; } catch { }
        }
        _lifetime.Dispose();
    }
}
