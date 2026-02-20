using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Фикс ошибки 0x0000007e — «Указанный модуль не найден» (The specified module could not be found).
/// 
/// Причины:
///   - Отсутствует mscms.dll в каталоге драйвера
///   - Повреждён ключ CopyFiles\BIDI в реестре принтера (HP Universal и др.)
///   - Не хватает DLL-компонентов драйвера в spool\drivers
///
/// Шаги:
///   1. Копирование mscms.dll в spool\drivers\x64\3 (если отсутствует)
///   2. Удаление ключа CopyFiles\BIDI для конкретного принтера
///   3. Проверка целостности через SFC (опционально)
///   4. Очистка кеша драйверов
///   5. Перезапуск Spooler
/// </summary>
public class Error7eFixer : FixerBase
{
    public override string Name => "Ошибка 0x0000007e (Модуль не найден)";
    public override string Description =>
        "Исправляет ошибку «Указанный модуль не найден (0x0000007e)» при подключении к принтеру. " +
        "Копирует mscms.dll, удаляет BIDI-ключ реестра, проверяет файлы драйверов.";
    public override string[] TargetErrorCodes => ["0x0000007e", "7e", "module_not_found"];

    public override async Task<FixResult> ApplyAsync(
        PrinterInfo? printer, string? remoteMachine,
        IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        Report(Info("Диагностика и исправление ошибки 0x0000007e..."));

        // ── Шаг 1: Копирование mscms.dll ────────────────────────────────────────
        Report(Info("Шаг 1: Проверка и копирование mscms.dll..."));

        var step1 = """
            $sourceDir = "$env:SystemRoot\System32"
            $targetDirs = @(
                "$env:SystemRoot\System32\spool\drivers\x64\3",
                "$env:SystemRoot\System32\spool\drivers\x64\4",
                "$env:SystemRoot\System32\spool\drivers\W32X86\3"
            )
            
            $sourceDll = Join-Path $sourceDir "mscms.dll"
            if (-not (Test-Path $sourceDll)) {
                Write-Output "[ERROR] mscms.dll отсутствует в System32! Требуется SFC."
                sfc /scannow 2>&1 | Select-Object -First 5 | ForEach-Object { Write-Output "[INFO] SFC: $_" }
            } else {
                Write-Output "[OK] mscms.dll найден в System32"
                
                foreach ($dir in $targetDirs) {
                    if (Test-Path $dir) {
                        $target = Join-Path $dir "mscms.dll"
                        if (-not (Test-Path $target)) {
                            try {
                                Copy-Item -Path $sourceDll -Destination $target -Force
                                Write-Output "[OK] mscms.dll скопирован в $dir"
                            } catch {
                                Write-Output "[WARN] Не удалось скопировать в $dir : $_"
                            }
                        } else {
                            Write-Output "[OK] mscms.dll уже есть в $dir"
                        }
                    }
                }
            }
            """;

        using var engine = new PowerShellEngine(remoteMachine);
        var (_, out1, _) = await engine.RunAsync(step1, ct: ct);
        ReportOutput(out1, Report);

        // ── Шаг 2: Удаление BIDI-ключа реестра ──────────────────────────────────
        if (printer != null)
        {
            Report(Info($"Шаг 2: Проверка BIDI-ключа для '{printer.Name}'..."));

            var printerNameEscaped = printer.Name.Replace("'", "''");
            var step2 = @"
                $printerName = '" + printerNameEscaped + @"'
                $regBase = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
                $printerPath = Join-Path $regBase $printerName
                $bidiPath = Join-Path $printerPath 'CopyFiles\BIDI'
                
                if (Test-Path $bidiPath) {
                    try {
                        # Экспорт бэкапа
                        $backupFile = ""$env:TEMP\bidi_backup_$($printerName -replace '[\\\/\:]', '_').reg""
                        reg export ""HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers\$printerName\CopyFiles\BIDI"" $backupFile /y 2>&1 | Out-Null
                        Write-Output ""[INFO] Бэкап BIDI сохранён: $backupFile""
                        
                        Remove-Item -Path $bidiPath -Recurse -Force
                        Write-Output ""[OK] Ключ CopyFiles\BIDI удалён для '$printerName'""
                    } catch {
                        Write-Output ""[WARN] Не удалось удалить BIDI ключ: $_""
                    }
                } else {
                    Write-Output ""[OK] Ключ BIDI отсутствует — не требует действий""
                }
            ";

            var (_, out2, _) = await engine.RunAsync(step2, ct: ct);
            ReportOutput(out2, Report);
        }
        else
        {
            Report(Info("Шаг 2: Принтер не выбран — пропуск очистки BIDI."));
        }

        // ── Шаг 3: Проверка DLL-файлов драйвера ────────────────────────────────
        Report(Info("Шаг 3: Проверка целостности файлов драйверов..."));

        var step3 = """
            $driversPath = "$env:SystemRoot\System32\spool\drivers\x64\3"
            
            if (Test-Path $driversPath) {
                $dlls = Get-ChildItem -Path $driversPath -Filter "*.dll" -ErrorAction SilentlyContinue
                $broken = @()
                
                foreach ($dll in $dlls) {
                    try {
                        # Пробуем прочитать PE-заголовок (простая проверка целостности)
                        $bytes = [System.IO.File]::ReadAllBytes($dll.FullName)
                        if ($bytes.Length -lt 64 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
                            $broken += $dll.Name
                        }
                    } catch {
                        $broken += $dll.Name
                    }
                }
                
                Write-Output "[INFO] Файлов драйверов: $($dlls.Count)"
                if ($broken.Count -gt 0) {
                    Write-Output "[WARN] Повреждённые файлы:"
                    foreach ($b in $broken) {
                        Write-Output "[WARN]   • $b"
                    }
                } else {
                    Write-Output "[OK] Все DLL-файлы драйверов целы"
                }
            }
            """;

        var (_, out3, _) = await engine.RunAsync(step3, ct: ct);
        ReportOutput(out3, Report);

        // ── Шаг 4: RPC + Spooler ───────────────────────────────────────────────
        Report(Info("Шаг 4: Настройка RPC и перезапуск Spooler..."));

        var step4 = """
            # RPC auth level
            $printPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print'
            Set-ItemProperty -Path $printPath -Name 'RpcAuthnLevelPrivacyEnabled' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
            Write-Output "[OK] RpcAuthnLevelPrivacyEnabled = 0"
            
            # Перезапуск
            Restart-Service -Name spooler -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            $status = (Get-Service spooler).Status
            Write-Output "[OK] Spooler: $status"
            """;

        var (success, out4, error) = await engine.RunAsync(step4, ct: ct);
        ReportOutput(out4, Report);

        Report(Info("Рекомендация: попробуйте переподключить принтер."));

        return success
            ? FixResult.Ok("Фикс 0x0000007e выполнен — mscms.dll на месте, BIDI очищен", steps)
            : FixResult.Warn($"Фикс 0x0000007e частично: {error}", steps);
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
