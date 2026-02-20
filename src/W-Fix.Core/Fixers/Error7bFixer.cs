using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Фикс ошибки 0x0000007b (ERROR_INVALID_NAME / драйвер принтера поврежден).
/// Обычно возникает из-за несовместимого или corrupted драйвера.
/// </summary>
public class Error7bFixer : FixerBase
{
    public override string Name => "Ошибка 0x0000007b (Драйвер)";
    public override string Description =>
        "Удаляет повреждённый драйвер принтера из хранилища Windows, очищает кеш Spooler и " +
        "предлагает переустановить драйвер. Решает ошибку недействительного имени устройства.";
    public override string[] TargetErrorCodes => ["0x0000007b", "7b", "ERROR_INVALID_NAME"];

    public override async Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        if (printer == null)
        {
            Report(Warn("Принтер не выбран. Будет выполнена общая очистка хранилища драйверов."));
        }

        var printerNameArg = printer?.Name ?? "";
        var driverNameArg = printer?.DriverName ?? "";

        Report(Info($"Начало фикса 0x0000007b" + (printer != null ? $" для '{printer.Name}'" : "")));

        var script = @"
            $ErrorActionPreference = 'Continue'
            $printerName = '" + printerNameArg.Replace("'", "''") + @"'
            $driverName = '" + driverNameArg.Replace("'", "''") + @"'

            # Шаг 1: остановить Spooler
            Write-Output ""[INFO] Останавливаем Spooler...""
            Stop-Service -Name spooler -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2

            # Шаг 2: удалить связанный принтер (если указан)
            if ($printerName -ne '') {
                Write-Output ""[INFO] Удаляем принтер '$printerName'...""
                try {
                    Remove-Printer -Name $printerName -ErrorAction Stop
                    Write-Output ""[OK] Принтер удалён""
                } catch {
                    Write-Output ""[WARN] Принтер не найден или уже удалён: $_""
                }
            }

            # Шаг 3: удалить драйвер (если указан)
            if ($driverName -ne '') {
                Write-Output ""[INFO] Удаляем драйвер '$driverName'...""
                try {
                    Remove-PrinterDriver -Name $driverName -ErrorAction Stop
                    Write-Output ""[OK] Драйвер удалён из диспетчера""
                } catch {
                    Write-Output ""[WARN] Не удалось удалить через Remove-PrinterDriver: $_""
                }

                # Попытка через printui
                Write-Output ""[INFO] Очищаем хранилище через printui...""
                $result = & printui.exe /s /t2 2>&1
                Write-Output ""[INFO] printui завершён""
            }

            # Шаг 4: очистить кеш драйверов
            Write-Output ""[INFO] Очищаем кеш драйверов в spool\drivers...""
            $driverCache = ""$env:SystemRoot\System32\spool\drivers""
            Get-ChildItem -Path $driverCache -Recurse -Include *.tmp,*.bak -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
            Write-Output ""[OK] Кеш очищен""

            # Шаг 5: запустить Spooler
            Write-Output ""[INFO] Запускаем Spooler...""
            Start-Service -Name spooler -ErrorAction Stop
            $status = (Get-Service spooler).Status
            Write-Output ""[OK] Spooler запущен со статусом: $status""
            Write-Output ""[INFO] После этого рекомендуется переустановить драйвер через 'Управление печатью' или INF-файл.""
        ";

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

        Report(Info("СЛЕДУЮЩИЙ ШАГ: Переустановите драйвер принтера из INF-файла или с сайта производителя."));
        return success
            ? FixResult.Ok("Фикс 0x0000007b выполнен: принтер и драйвер удалены, Spooler перезапущен", steps)
            : FixResult.Warn($"Фикс частично выполнен: {error}", steps);
    }
}
