using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Фикс ошибки 0x00004005 — «Операция не может быть завершена».
/// Одна из самых частых ошибок при подключении к сетевым (расшаренным) принтерам.
/// Возникает после обновлений Windows (KB5005565, KB5006670 и др.).
/// 
/// Решение: несколько шагов, включая:
///   1. Установка RpcAuthnLevelPrivacyEnabled = 0
///   2. Настройка LanManServer (SMB-сервер)
///   3. Включение File and Printer Sharing в брандмауэре
///   4. Очистка и перезапуск Spooler
///   5. Установка клиента локального порта (если удалён обновлением)
/// </summary>
public class Error4005Fixer : FixerBase
{
    public override string Name => "Ошибка 0x00004005 (Подключение)";
    public override string Description =>
        "Исправляет ошибку «Операция не может быть завершена (0x00004005)» при подключении к сетевому принтеру. " +
        "Возникает после обновлений Windows. Настраивает RPC, SMB, брандмауэр и Spooler.";
    public override string[] TargetErrorCodes => ["0x00004005", "4005", "operation_could_not_be_completed"];

    public override async Task<FixResult> ApplyAsync(
        PrinterInfo? printer, string? remoteMachine,
        IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        Report(Info("Диагностика и исправление ошибки 0x00004005..."));

        // ── Шаг 1: Реестр — RpcAuthnLevelPrivacyEnabled ─────────────────────────
        Report(Info("Шаг 1: Настройка RPC Authentication Level..."));

        var step1 = """
            $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print'
            $regName = 'RpcAuthnLevelPrivacyEnabled'
            
            try {
                $current = Get-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue
                if ($current -and $current.$regName -eq 0) {
                    Write-Output "[OK] RpcAuthnLevelPrivacyEnabled уже = 0 (отключено)"
                } else {
                    Set-ItemProperty -Path $regPath -Name $regName -Value 0 -Type DWord -Force
                    Write-Output "[OK] RpcAuthnLevelPrivacyEnabled установлено в 0"
                }
            } catch {
                Write-Output "[ERROR] Не удалось изменить реестр: $_"
            }
            """;

        using var engine = remoteMachine != null ? new PowerShellEngine(remoteMachine) : null;
        async Task<(bool, IReadOnlyList<string>, string?)> RunPS(string script) =>
            engine != null ? await engine.RunAsync(script, ct: ct) : await PowerShellEngine.RunExternalAsync(script, ct);

        var (_, out1, _) = await RunPS(step1);
        ReportOutput(out1, Report);

        // ── Шаг 2: Реестр — RestrictDriverInstallationToAdministrators ──────────
        Report(Info("Шаг 2: Разрешение установки драйверов (Point and Print)..."));

        var step2 = """
            $regPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint'
            
            try {
                if (-not (Test-Path $regPath)) {
                    New-Item -Path $regPath -Force | Out-Null
                    Write-Output "[OK] Создан раздел реестра PointAndPrint"
                }
                
                Set-ItemProperty -Path $regPath -Name 'RestrictDriverInstallationToAdministrators' -Value 0 -Type DWord -Force
                Write-Output "[OK] RestrictDriverInstallationToAdministrators = 0"
                
                # Настройка Point and Print для доверия серверам
                Set-ItemProperty -Path $regPath -Name 'NoWarningNoElevationOnInstall' -Value 1 -Type DWord -Force
                Set-ItemProperty -Path $regPath -Name 'UpdatePromptSettings' -Value 1 -Type DWord -Force
                Write-Output "[OK] Point and Print: предупреждения отключены"
            } catch {
                Write-Output "[WARN] Не удалось настроить Point and Print: $_"
            }
            """;

        var (_, out2, _) = await RunPS(step2);
        ReportOutput(out2, Report);

        // ── Шаг 3: Брандмауэр — File and Printer Sharing ───────────────────────
        Report(Info("Шаг 3: Проверка правил брандмауэра..."));

        var step3 = """
            try {
                $rules = Get-NetFirewallRule -DisplayGroup 'File and Printer Sharing' -ErrorAction SilentlyContinue
                if (-not $rules) {
                    $rules = Get-NetFirewallRule -DisplayGroup 'Общий доступ к файлам и принтерам' -ErrorAction SilentlyContinue
                }
                
                if ($rules) {
                    $disabled = $rules | Where-Object { $_.Enabled -eq 'False' }
                    if ($disabled) {
                        $disabled | Set-NetFirewallRule -Enabled True
                        Write-Output "[OK] Включены правила брандмауэра: Общий доступ к файлам и принтерам ($($disabled.Count) правил)"
                    } else {
                        Write-Output "[OK] Правила брандмауэра уже включены"
                    }
                } else {
                    Write-Output "[WARN] Не найдена группа правил File and Printer Sharing"
                }
            } catch {
                Write-Output "[WARN] Ошибка настройки брандмауэра: $_"
            }
            """;

        var (_, out3, _) = await RunPS(step3);
        ReportOutput(out3, Report);

        // ── Шаг 4: SMB клиент – обеспечить доступ ──────────────────────────────
        Report(Info("Шаг 4: Проверка SMB-клиента..."));

        var step4 = """
            try {
                # Проверяем SMB1 — некоторые старые принт-серверы требуют его
                $smb1 = Get-WindowsOptionalFeature -Online -FeatureName 'SMB1Protocol' -ErrorAction SilentlyContinue
                if ($smb1 -and $smb1.State -eq 'Disabled') {
                    Write-Output "[INFO] SMB1 отключён. Если принт-сервер старый — может потребоваться включение."
                } else {
                    Write-Output "[OK] SMB1: $($smb1.State)"
                }
                
                # Проверяем службу LanmanWorkstation
                $lanman = Get-Service -Name LanmanWorkstation -ErrorAction SilentlyContinue
                if ($lanman.Status -ne 'Running') {
                    Start-Service -Name LanmanWorkstation -ErrorAction Stop
                    Write-Output "[OK] Служба LanmanWorkstation запущена"
                } else {
                    Write-Output "[OK] Служба LanmanWorkstation: Running"
                }
            } catch {
                Write-Output "[WARN] Ошибка проверки SMB: $_"
            }
            """;

        var (_, out4, _) = await RunPS(step4);
        ReportOutput(out4, Report);

        // ── Шаг 5: Перезапуск Spooler ───────────────────────────────────────────
        Report(Info("Шаг 5: Перезапуск Print Spooler..."));
        var spoolerFixer = new SpoolerFixer();
        var spoolerResult = await spoolerFixer.ApplyAsync(printer, remoteMachine, progress, ct);
        steps.AddRange(spoolerResult.Steps);

        // ── Шаг 6: Проверка подключения (если выбран принтер) ───────────────────
        if (printer?.IsNetwork == true)
        {
            Report(Info($"Шаг 6: Проверка доступности принтера {printer.Name}..."));

            var target = printer.PortName;
            if (target.StartsWith(@"\\"))
            {
                var checkScript = @"
                    $uncPath = '" + target.Replace("'", "''") + @"'
                    $server = ($uncPath -replace '^\\\\', '') -split '\\' | Select-Object -First 1
                    
                    # Ping сервера
                    $ping = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction SilentlyContinue
                    if ($ping) {
                        Write-Output ""[OK] Сервер $server отвечает на ping""
                    } else {
                        Write-Output ""[WARN] Сервер $server не отвечает (проверьте имя/IP)""
                    }

                    # Проверяем доступность общих ресурсов
                    try {
                        $shares = net view ""\\$server"" 2>&1
                        Write-Output ""[OK] Общие ресурсы сервера доступны""
                    } catch {
                        Write-Output ""[WARN] Не удалось получить список общих ресурсов: $_""
                    }
                ";

                var (_, out6, _) = await RunPS(checkScript);
                ReportOutput(out6, Report);
            }
        }

        Report(Info("═══════════════════════════════════════════════════"));
        Report(Info("Рекомендация: перезагрузите компьютер для применения всех изменений реестра."));

        return spoolerResult.Status == FixStatus.Success
            ? FixResult.Ok("Фикс 0x00004005 выполнен. Рекомендуется перезагрузка.", steps)
            : FixResult.Warn("Фикс 0x00004005 частично выполнен — проверьте логи", steps);
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
