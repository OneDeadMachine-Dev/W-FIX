using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Универсальный сброс Print Spooler + очистка очереди.
/// Используется как самостоятельный фикс и как компонент в других фиксерах.
/// </summary>
public class SpoolerFixer : FixerBase
{
    public override string Name => "Перезапуск Print Spooler";
    public override string Description =>
        "Останавливает службу Spooler, очищает все задания в папке PRINTERS, запускает службу заново. " +
        "Решает зависшие очереди печати, ошибку 0x00000008 (insufficient memory) и общие сбои Spooler.";
    public override string[] TargetErrorCodes => ["0x00000008", "0x00000006", "spooler"];

    public override async Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        Report(Info("Остановка службы Print Spooler..."));

        var script = """
            $ErrorActionPreference = 'Stop'
            try {
                Stop-Service -Name spooler -Force -ErrorAction Stop
                Write-Output "[OK] Служба Spooler остановлена"
            } catch {
                Write-Output "[WARN] Не удалось остановить Spooler: $_"
            }

            $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
            $files = Get-ChildItem -Path $spoolPath -Include *.SHD,*.SPL -Recurse -ErrorAction SilentlyContinue
            $count = $files.Count
            $files | Remove-Item -Force -ErrorAction SilentlyContinue
            Write-Output "[OK] Удалено файлов очереди: $count"

            try {
                Start-Service -Name spooler -ErrorAction Stop
                Write-Output "[OK] Служба Spooler запущена"
                $status = (Get-Service spooler).Status
                Write-Output "[OK] Статус: $status"
            } catch {
                Write-Output "[ERROR] Не удалось запустить Spooler: $_"
                exit 1
            }
            """;

        using var engine = new PowerShellEngine(remoteMachine);
        var (success, output, error) = await engine.RunAsync(script, ct: ct);

        foreach (var line in output)
        {
            var level = line.StartsWith("[OK]") ? Models.LogLevel.Success
                      : line.StartsWith("[WARN]") ? Models.LogLevel.Warning
                      : line.StartsWith("[ERROR]") ? Models.LogLevel.Error
                      : Models.LogLevel.Info;
            Report(new LogEntry(level, line));
        }

        return success
            ? FixResult.Ok("Print Spooler успешно перезапущен, очередь очищена", steps)
            : FixResult.Fail($"Ошибка при перезапуске Spooler: {error}", steps);
    }
}
