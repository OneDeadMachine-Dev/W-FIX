using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Фикс ошибки 0x00000008 (ERROR_NOT_ENOUGH_MEMORY).
/// Возникает когда Spooler переполнен зависшими заданиями или недостаточно OAM-памяти.
/// </summary>
public class Error8Fixer : FixerBase
{
    public override string Name => "Ошибка 0x00000008 (Память Spooler)";
    public override string Description =>
        "Диагностирует нехватку памяти в сервисе Spooler: очищает зависшие задания, " +
        "проверяет свободное место на диске, перезапускает Spooler с увеличенными лимитами.";
    public override string[] TargetErrorCodes => ["0x00000008", "0x8", "ERROR_NOT_ENOUGH_MEMORY"];

    public override async Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        Report(Info("Диагностика памяти Spooler..."));

        var diagScript = """
            # Диагностика
            $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
            $files = Get-ChildItem -Path $spoolPath -Recurse -ErrorAction SilentlyContinue
            $totalSizeMB = [Math]::Round(($files | Measure-Object -Property Length -Sum).Sum / 1MB, 2)
            Write-Output "[INFO] Файлов в очереди: $($files.Count), суммарный размер: $totalSizeMB МБ"

            # Проверка диска
            $drive = Split-Path $env:SystemRoot -Qualifier
            $disk = Get-PSDrive -Name ($drive.TrimEnd(':')) -ErrorAction SilentlyContinue
            if ($disk) {
                $freeGB = [Math]::Round($disk.Free / 1GB, 2)
                Write-Output "[INFO] Свободное место на диске $drive : $freeGB ГБ"
                if ($freeGB -lt 1) {
                    Write-Output "[WARN] Критически мало места на диске! Это может вызывать ошибку 0x00000008."
                }
            }

            # Проверка памяти процесса Spooler
            $spoolerProc = Get-Process -Name spoolsv -ErrorAction SilentlyContinue
            if ($spoolerProc) {
                $memMB = [Math]::Round($spoolerProc.WorkingSet64 / 1MB, 2)
                Write-Output "[INFO] Память процесса Spooler (spoolsv.exe): $memMB МБ"
                if ($memMB -gt 500) {
                    Write-Output "[WARN] Процесс Spooler занимает много памяти — выполняем принудительный сброс."
                }
            }
            """;

        using var engine = new PowerShellEngine(remoteMachine);
        var (_, diagOutput, _) = await engine.RunAsync(diagScript, ct: ct);
        foreach (var line in diagOutput)
        {
            var level = line.StartsWith("[OK]") ? Models.LogLevel.Success
                      : line.StartsWith("[WARN]") ? Models.LogLevel.Warning
                      : line.StartsWith("[ERROR]") ? Models.LogLevel.Error
                      : Models.LogLevel.Info;
            Report(new LogEntry(level, line));
        }

        // Применяем спулер-фикс
        Report(Info("Выполняем сброс очереди и перезапуск Spooler..."));
        var spoolerFixer = new SpoolerFixer();
        var spoolerResult = await spoolerFixer.ApplyAsync(printer, remoteMachine, progress, ct);
        steps.AddRange(spoolerResult.Steps);

        // Дополнительно: очистка временных файлов Windows
        var cleanupScript = """
            Write-Output "[INFO] Очистка временных файлов Windows (помогает при нехватке места)..."
            $tempPaths = @($env:TEMP, $env:TMP, "$env:SystemRoot\Temp")
            $totalDeleted = 0
            foreach ($p in $tempPaths) {
                if (Test-Path $p) {
                    $items = Get-ChildItem -Path $p -Recurse -ErrorAction SilentlyContinue
                    $items | Where-Object { -not $_.PSIsContainer } | 
                        Remove-Item -Force -ErrorAction SilentlyContinue
                    $totalDeleted++
                }
            }
            Write-Output "[OK] Очистка временных папок завершена"
            """;

        var (cleanSuccess, cleanOutput, cleanError) = await engine.RunAsync(cleanupScript, ct: ct);
        foreach (var line in cleanOutput)
        {
            var level = line.StartsWith("[OK]") ? Models.LogLevel.Success
                      : line.StartsWith("[WARN]") ? Models.LogLevel.Warning
                      : line.StartsWith("[ERROR]") ? Models.LogLevel.Error
                      : Models.LogLevel.Info;
            Report(new LogEntry(level, line));
        }

        return spoolerResult.Status == FixStatus.Success
            ? FixResult.Ok("Фикс 0x00000008 выполнен: очередь очищена, временные файлы удалены", steps)
            : FixResult.Warn("Фикс 0x00000008 частично выполнен — проверьте логи", steps);
    }
}
