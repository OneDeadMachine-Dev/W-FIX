using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Фикс ошибки 0x00000002 — «Системе не удаётся найти указанный файл»
/// при подключении к сетевому/расшаренному принтеру.
///
/// Причины:
///   - Засорены каталоги Print Processors (папки 1-499 в prtprocs\x64)
///   - PendingFileRenameOperations в реестре блокирует драйвер
///   - Повреждены файлы драйвера на print-сервере
///   - RestrictDriverInstallationToAdministrators блокирует установку
///
/// Шаги:
///   1. Очистка лишних папок в Print Processors (оставляем только winprint.dll)
///   2. Удаление PendingFileRenameOperations (если содержит принтерные записи)
///   3. Настройка Point and Print политик
///   4. Переустановка драйвера через pnputil (если возможно)
///   5. Перезапуск Spooler
/// </summary>
public class Error02Fixer : FixerBase
{
    public override string Name => "Ошибка 0x00000002 (Файл не найден)";
    public override string Description =>
        "Исправляет ошибку «Системе не удаётся найти указанный файл (0x00000002)» при подключении к " +
        "сетевому принтеру. Чистит Print Processors, реестр PendingFileRename и настраивает Point and Print.";
    public override string[] TargetErrorCodes => ["0x00000002", "0x00000003", "file_not_found", "system_cannot_find"];

    public override async Task<FixResult> ApplyAsync(
        PrinterInfo? printer, string? remoteMachine,
        IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        Report(Info("Диагностика и исправление ошибки 0x00000002..."));

        // ── Шаг 1: Очистка Print Processors ─────────────────────────────────────
        Report(Info("Шаг 1: Очистка лишних папок Print Processors..."));

        var step1 = """
            $prtProcsPath = "$env:SystemRoot\System32\spool\prtprocs\x64"
            
            if (Test-Path $prtProcsPath) {
                # Получаем все подпапки с числовыми именами (1, 2, ... 499)
                $numericDirs = Get-ChildItem -Path $prtProcsPath -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '^\d+$' }
                
                if ($numericDirs.Count -gt 0) {
                    Write-Output "[INFO] Найдено $($numericDirs.Count) числовых папок в prtprocs\x64"
                    
                    # Останавливаем Spooler для безопасного удаления
                    Stop-Service -Name spooler -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 1
                    
                    $removed = 0
                    foreach ($dir in $numericDirs) {
                        try {
                            Remove-Item -Path $dir.FullName -Recurse -Force -ErrorAction Stop
                            $removed++
                        } catch {
                            Write-Output "[WARN] Не удалось удалить: $($dir.Name) — $_"
                        }
                    }
                    Write-Output "[OK] Удалено числовых папок: $removed из $($numericDirs.Count)"
                    
                    Start-Service -Name spooler -ErrorAction SilentlyContinue
                } else {
                    Write-Output "[OK] Лишних папок в prtprocs\x64 нет"
                }
                
                # Проверяем наличие winprint.dll
                $winprint = Join-Path $prtProcsPath "winprint.dll"
                if (Test-Path $winprint) {
                    Write-Output "[OK] winprint.dll на месте"
                } else {
                    Write-Output "[WARN] winprint.dll отсутствует! Запускаем sfc..."
                    sfc /scannow 2>&1 | Out-Null
                    Write-Output "[INFO] SFC завершён"
                }
            } else {
                Write-Output "[WARN] Каталог $prtProcsPath не найден"
            }
            """;

        using var engine = new PowerShellEngine(remoteMachine);
        var (_, out1, _) = await engine.RunAsync(step1, ct: ct);
        ReportOutput(out1, Report);

        // ── Шаг 2: Очистка Print Environments в реестре ─────────────────────────
        Report(Info("Шаг 2: Проверка Print Environments в реестре..."));

        var step2 = """
            $envPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Print Processors'
            
            try {
                $processors = Get-ChildItem -Path $envPath -ErrorAction SilentlyContinue
                $toRemove = $processors | Where-Object { $_.PSChildName -ne 'winprint' }
                
                if ($toRemove.Count -gt 0) {
                    Write-Output "[INFO] Найдено $($toRemove.Count) сторонних Print Processor(ов)"
                    foreach ($proc in $toRemove) {
                        $dllPath = (Get-ItemProperty -Path $proc.PSPath -Name 'Driver' -ErrorAction SilentlyContinue).Driver
                        Write-Output "[INFO]   • $($proc.PSChildName) → $dllPath"
                    }
                    Write-Output "[INFO] Сторонние процессоры оставлены (удалите вручную если нужно)"
                } else {
                    Write-Output "[OK] Только winprint — всё чисто"
                }
            } catch {
                Write-Output "[WARN] Ошибка проверки реестра: $_"
            }
            """;

        var (_, out2, _) = await engine.RunAsync(step2, ct: ct);
        ReportOutput(out2, Report);

        // ── Шаг 3: PendingFileRenameOperations ──────────────────────────────────
        Report(Info("Шаг 3: Проверка PendingFileRenameOperations..."));

        var step3 = """
            $smPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
            
            try {
                $pending = Get-ItemProperty -Path $smPath -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
                
                if ($pending) {
                    $ops = $pending.PendingFileRenameOperations
                    $printerRelated = $ops | Where-Object { $_ -like '*spool*' -or $_ -like '*print*' -or $_ -like '*driver*' }
                    
                    if ($printerRelated.Count -gt 0) {
                        Write-Output "[WARN] Найдены print-related PendingFileRename записи: $($printerRelated.Count)"
                        # Фильтруем — оставляем только не-принтерные
                        $clean = $ops | Where-Object { $_ -notlike '*spool*' -and $_ -notlike '*print*' }
                        if ($clean.Count -gt 0) {
                            Set-ItemProperty -Path $smPath -Name 'PendingFileRenameOperations' -Value $clean -Force
                        } else {
                            Remove-ItemProperty -Path $smPath -Name 'PendingFileRenameOperations' -Force
                        }
                        Write-Output "[OK] Принтерные PendingFileRename записи удалены"
                    } else {
                        Write-Output "[OK] PendingFileRenameOperations не содержит принтерных записей"
                    }
                } else {
                    Write-Output "[OK] PendingFileRenameOperations отсутствует"
                }
            } catch {
                Write-Output "[WARN] Ошибка: $_"
            }
            """;

        var (_, out3, _) = await engine.RunAsync(step3, ct: ct);
        ReportOutput(out3, Report);

        // ── Шаг 4: Point and Print + RPC ────────────────────────────────────────
        Report(Info("Шаг 4: Настройка Point and Print..."));

        var step4 = """
            try {
                $ppPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint'
                if (-not (Test-Path $ppPath)) {
                    New-Item -Path $ppPath -Force | Out-Null
                }
                Set-ItemProperty -Path $ppPath -Name 'RestrictDriverInstallationToAdministrators' -Value 0 -Type DWord -Force
                Set-ItemProperty -Path $ppPath -Name 'NoWarningNoElevationOnInstall' -Value 1 -Type DWord -Force
                Write-Output "[OK] Point and Print: разрешена установка драйверов"
                
                # RPC
                $printPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print'
                Set-ItemProperty -Path $printPath -Name 'RpcAuthnLevelPrivacyEnabled' -Value 0 -Type DWord -Force
                Write-Output "[OK] RpcAuthnLevelPrivacyEnabled = 0"
            } catch {
                Write-Output "[WARN] Ошибка настройки: $_"
            }
            """;

        var (_, out4, _) = await engine.RunAsync(step4, ct: ct);
        ReportOutput(out4, Report);

        // ── Шаг 5: Перезапуск Spooler ───────────────────────────────────────────
        Report(Info("Шаг 5: Перезапуск Print Spooler..."));
        var spooler = new SpoolerFixer();
        var spoolerResult = await spooler.ApplyAsync(printer, remoteMachine, progress, ct);
        steps.AddRange(spoolerResult.Steps);

        Report(Info("Рекомендация: перезагрузите компьютер, затем повторите подключение к принтеру."));

        return spoolerResult.Status == FixStatus.Success
            ? FixResult.Ok("Фикс 0x00000002 выполнен", steps)
            : FixResult.Warn("Фикс 0x00000002 частично выполнен", steps);
    }

    private static void ReportOutput(IReadOnlyList<string> output, Action<LogEntry> report)
    {
        foreach (var line in output)
        {
            var level = line.StartsWith("[OK]") ? Models.LogLevel.Success
                      : line.StartsWith("[WARN]") ? Models.LogLevel.Warning
                      : line.StartsWith("[ERROR]") ? Models.LogLevel.Error
                      : Models.LogLevel.Info;
            report(new LogEntry(level, line));
        }
    }
}
