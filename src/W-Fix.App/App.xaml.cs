using System.IO;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
using WFix.App.Services;
using WFix.App.ViewModels;
using WFix.Core.Abstractions;
using WFix.Core.Catalog;
using WFix.Core.Diagnostics;
using WFix.Core.Infrastructure;
using WFix.Core.Pairing;
using WFix.Core.Remote;
using WFix.Core.Repair;
using WFix.Core.Reporting;
using WFix.Core.Services;

namespace WFix.App;

public partial class App : Application
{
    public IServiceProvider Services { get; private set; } = null!;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Инициализация Serilog: лог в файл рядом с exe
        var logDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "W-Fix", "Logs");
        Directory.CreateDirectory(logDir);

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(Path.Combine(logDir, "w-fix-.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 14)
#if DEBUG
            .WriteTo.Debug()
#endif
            .CreateLogger();

        Log.Information("=== W-Fix запущен | Пользователь: {User} | Машина: {Machine} ===",
            Environment.UserName, Environment.MachineName);

        // Перехват необработанных исключений
        DispatcherUnhandledException += (_, ex) =>
        {
            Log.Fatal(ex.Exception, "Необработанное исключение в UI");
            MessageBox.Show(
                $"Критическая ошибка:\n{ex.Exception.Message}\n\nПодробности в лог-файле:\n{logDir}",
                "W-Fix — Критическая ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
            ex.Handled = true;
        };

        Services = ConfigureServices();
        Services.GetRequiredService<MainWindow>().Show();
    }

    public RemoteCenterWindow CreateRemoteCenterWindow() =>
        Services.GetRequiredService<RemoteCenterWindow>();

    public PairRepairWindow CreatePairRepairWindow() =>
        Services.GetRequiredService<PairRepairWindow>();

    protected override void OnExit(ExitEventArgs e)
    {
        Log.Information("W-Fix завершён");
        (Services as IDisposable)?.Dispose();
        Log.CloseAndFlush();
        base.OnExit(e);
    }

    private static IServiceProvider ConfigureServices()
    {
        var services = new ServiceCollection();

        services.AddSingleton<WmiService>();
        services.AddSingleton<ActiveDirectoryService>();
        services.AddSingleton<FixerRegistry>();
        services.AddSingleton<SystemStateBackupService>();
        services.AddSingleton<InteractiveUserPowerShellService>();

        services.AddSingleton<ICredentialStore, WindowsCredentialStore>();
        services.AddSingleton<IRemoteSessionFactory, PowerShellRemoteSessionFactory>();
        services.AddSingleton<IRemotePreflightService, RemotePreflightService>();
        services.AddSingleton<IPrinterInventoryService, RemotePrinterInventoryService>();
        services.AddSingleton<IRemoteMaintenanceService, RemoteMaintenanceService>();
        services.AddSingleton<IRepairActionRegistry, RepairActionRegistry>();
        services.AddSingleton<IKnownIssueCatalog, KnownIssueCatalogService>();
        services.AddSingleton<IRepairPlanner, RepairPlanner>();
        services.AddSingleton<IRunReportService, RunReportService>();
        services.AddSingleton<ISupportBundleService, SupportBundleService>();

        services.AddSingleton<IDiagnosticRule, SpoolerDiagnosticRule>();
        services.AddSingleton<IDiagnosticRule, PrinterStateDiagnosticRule>();
        services.AddSingleton<IDiagnosticRule, DefaultPrinterDiagnosticRule>();
        services.AddSingleton<IDiagnosticRule, ProtectedPrintModeDiagnosticRule>();
        services.AddSingleton<IDiagnosticRule, UsbIppUpdateDiagnosticRule>();
        services.AddSingleton<IDiagnosticService, DiagnosticService>();
        services.AddSingleton<IRepairExecutor, RepairExecutor>();

        services.AddSingleton<IPairInvitationValidator, PairInvitationValidator>();
        services.AddSingleton<IPairSessionTransport, TlsPairSessionTransport>();
        services.AddSingleton<IPairFirewallLeaseService, WindowsPairFirewallLeaseService>();
        services.AddSingleton<IPairFileService, PairFileService>();
        services.AddSingleton<IPairInventoryService, PairInventoryService>();
        services.AddSingleton<IPairDiagnosticRule, PairDiscoveryDiagnosticRule>();
        services.AddSingleton<IPairDiagnosticRule, PairSmbDiagnosticRule>();
        services.AddSingleton<IPairDiagnosticRule, PairRpcDiagnosticRule>();
        services.AddSingleton<IPairDiagnosticRule, PairPrinterShareDiagnosticRule>();
        services.AddSingleton<IPairDiagnosticRule, PairRpcCompatibilityDiagnosticRule>();
        services.AddSingleton<IPairDiagnosticService, PairDiagnosticService>();
        services.AddSingleton<IPairRepairPlanner, PairRepairPlanner>();
        services.AddSingleton<IPairRepairActionRegistry, PairRepairActionRegistry>();
        services.AddSingleton<IPairActionDispatcher, RegistryPairActionDispatcher>();
        services.AddTransient<IPairAgentCommandLoop, PairAgentCommandLoop>();
        services.AddSingleton<IPairRunReportService, PairRunReportService>();
        services.AddSingleton<INetworkCredentialProvisioner, WindowsNetworkCredentialProvisioner>();

        services.AddSingleton<IUserPromptService, UserPromptService>();
        services.AddSingleton<MainWindowViewModel>();
        services.AddTransient<RemoteCenterViewModel>();
        services.AddTransient<PairRepairViewModel>();
        services.AddSingleton<MainWindow>();
        services.AddTransient<RemoteCenterWindow>();
        services.AddTransient<PairRepairWindow>();

        return services.BuildServiceProvider(new ServiceProviderOptions
        {
            ValidateOnBuild = true,
            ValidateScopes = true
        });
    }
}
