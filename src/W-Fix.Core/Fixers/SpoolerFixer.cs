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
        "Останавливает службу Spooler, очищает все задания (локальные файлы + сетевые задания через WMI), " +
        "запускает службу заново. Решает зависшие очереди печати и общие сбои Spooler.";
    public override string[] TargetErrorCodes => ["0x00000008", "0x00000006", "spooler"];

    public override async Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        Report(Info("Остановка службы Print Spooler..."));

        // Используем внешний powershell.exe — нужен CimCmdlets и Remove-PrintJob (Desktop-only)
        var script = @"
            $ErrorActionPreference = 'Continue'

            # Шаг 1: Очищаем задания через WMI перед остановкой (сетевые принтеры Win10 22H2+)
            try {
                $jobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue
                if ($jobs) {
                    $count = ($jobs | Measure-Object).Count
                    $jobs | Remove-CimInstance -ErrorAction SilentlyContinue
                    Write-Output ""[OK] Удалено заданий через WMI: $count""
                } else {
                    Write-Output ""[INFO] Активных заданий WMI не найдено""
                }
            } catch {
                Write-Output ""[WARN] Не удалось очистить задания через WMI: $_""
            }

            # Шаг 2: Остановка Spooler
            try {
                Stop-Service -Name spooler -Force -ErrorAction Stop
                Start-Sleep -Seconds 1
                Write-Output ""[OK] Служба Spooler остановлена""
            } catch {
                # Принудительное завершение процесса если Stop-Service не помог
                Write-Output ""[WARN] Stop-Service не сработал, завершаем процесс: $_""
                Get-Process -Name spoolsv -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
            }

            # Шаг 3: Удаление файлов очереди (локальные принтеры)
            $spoolPath = ""$env:SystemRoot\System32\spool\PRINTERS""
            $files = Get-ChildItem -Path $spoolPath -Include *.SHD,*.SPL -Recurse -ErrorAction SilentlyContinue
            $fileCount = ($files | Measure-Object).Count
            $files | Remove-Item -Force -ErrorAction SilentlyContinue
            Write-Output ""[OK] Удалено файлов очереди: $fileCount""

            # Шаг 4: Запуск Spooler
            try {
                Start-Service -Name spooler -ErrorAction Stop
                $status = (Get-Service spooler).Status
                Write-Output ""[OK] Служба Spooler запущена: $status""
            } catch {
                Write-Output ""[ERROR] Не удалось запустить Spooler: $_""
                exit 1
            }
        ";

        // RunExternalAsync = powershell.exe (Desktop) — есть CimCmdlets и Remove-PrintJob
        var (success, output, error) = remoteMachine == null
            ? await PowerShellEngine.RunExternalAsync(script, ct)
            : await new PowerShellEngine(remoteMachine).RunAsync(script, ct: ct);

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


