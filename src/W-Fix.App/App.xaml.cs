using System.IO;
using System.Windows;
using Serilog;

namespace WFix.App;

public partial class App : Application
{
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
    }

    protected override void OnExit(ExitEventArgs e)
    {
        Log.Information("W-Fix завершён");
        Log.CloseAndFlush();
        base.OnExit(e);
    }
}
