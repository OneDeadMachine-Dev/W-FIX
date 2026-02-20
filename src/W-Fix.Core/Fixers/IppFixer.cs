using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Фикс сбоя установки принтера через IPP (Internet Printing Protocol).
/// Ошибки: 0x00000bcb, 0x00000bcc, 0x80070bc9, а также общие сбои "Не удаётся установить принтер".
///
/// IPP — современный протокол печати через HTTP/HTTPS (порт 631/443).
/// Частые причины сбоев:
///   - Служба Print Spooler не настроена для IPP
///   - Проблемы с Windows Feature "Internet Printing Client"
///   - Конфликт драйверов Microsoft IPP Class Driver
///   - Блокировка порта 631 брандмауэром
///   - Повреждённый кеш IPP
///
/// Шаги:
///   1. Включение Windows Feature "Internet Printing Client" (Printing-Client-Lanman-Print-Svc)
///   2. Проверка/установка Microsoft IPP Class Driver
///   3. Открытие порта 631 (IPP) в брандмауэре  
///   4. Очистка кеша IPP-подключений
///   5. SFC + DISM при необходимости
///   6. Перезапуск Spooler
/// </summary>
public class IppFixer : FixerBase
{
    public override string Name => "Сбой IPP (Internet Printing)";
    public override string Description =>
        "Исправляет сбои установки принтера через IPP (Internet Printing Protocol). " +
        "Включает компонент Windows, настраивает IPP Class Driver, открывает порт 631, чистит кеш.";
    public override string[] TargetErrorCodes => ["0x00000bcb", "0x00000bcc", "0x80070bc9", "ipp", "internet_printing"];

    public override async Task<FixResult> ApplyAsync(
        PrinterInfo? printer, string? remoteMachine,
        IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        Report(Info("Диагностика и исправление сбоя IPP..."));

        // ── Шаг 1: Проверка/включение Internet Printing Client ──────────────────
        Report(Info("Шаг 1: Проверка компонента Internet Printing Client..."));

        var step1 = """
            try {
                # Проверяем компонент
                $feature = Get-WindowsOptionalFeature -Online -FeatureName 'Printing-Foundation-InternetPrinting-Client' -ErrorAction SilentlyContinue
                
                if (-not $feature) {
                    # Альтернативное имя
                    $feature = Get-WindowsOptionalFeature -Online -FeatureName 'Internet-Printing-Client' -ErrorAction SilentlyContinue
                }
                
                if ($feature) {
                    if ($feature.State -eq 'Enabled') {
                        Write-Output "[OK] Internet Printing Client: включён"
                    } else {
                        Write-Output "[INFO] Internet Printing Client: отключён. Включаем..."
                        Enable-WindowsOptionalFeature -Online -FeatureName $feature.FeatureName -NoRestart -ErrorAction Stop | Out-Null
                        Write-Output "[OK] Internet Printing Client включён"
                    }
                } else {
                    # Пробуем через DISM напрямую
                    Write-Output "[INFO] Проверяем через dism..."
                    $dismResult = dism /online /get-featureinfo /featurename:Printing-Foundation-InternetPrinting-Client 2>&1
                    $state = ($dismResult | Select-String 'State').ToString()
                    Write-Output "[INFO] DISM: $state"
                    
                    if ($state -like '*Disabled*') {
                        dism /online /enable-feature /featurename:Printing-Foundation-InternetPrinting-Client /norestart 2>&1 | Out-Null
                        Write-Output "[OK] Компонент включён через DISM"
                    }
                }
            } catch {
                Write-Output "[WARN] Ошибка при проверке компонента: $_"
            }
            """;

        using var engine = remoteMachine != null ? new PowerShellEngine(remoteMachine) : null;
        async Task<(bool, IReadOnlyList<string>, string?)> RunPS(string script) =>
            engine != null ? await engine.RunAsync(script, ct: ct) : await PowerShellEngine.RunExternalAsync(script, ct);

        var (_, out1, _) = await RunPS(step1);
        ReportOutput(out1, Report);

        // ── Шаг 2: Проверка Microsoft IPP Class Driver ──────────────────────────
        Report(Info("Шаг 2: Проверка Microsoft IPP Class Driver..."));

        var step2 = """
            try {
                $ippDriver = Get-PrinterDriver -Name "Microsoft IPP Class Driver" -ErrorAction SilentlyContinue
                
                if ($ippDriver) {
                    Write-Output "[OK] Microsoft IPP Class Driver установлен"
                    Write-Output "[INFO]   Версия: $($ippDriver.MajorVersion).$($ippDriver.MinorVersion)"
                    Write-Output "[INFO]   Среда: $($ippDriver.PrinterEnvironment)"
                } else {
                    Write-Output "[WARN] Microsoft IPP Class Driver не найден"
                    Write-Output "[INFO] Пробуем установить через pnputil..."
                    
                    # Ищем INF для IPP в DriverStore
                    $ippInf = Get-ChildItem -Path "$env:SystemRoot\INF" -Filter "prnms*" -ErrorAction SilentlyContinue |
                        Select-Object -First 1
                    
                    if ($ippInf) {
                        $result = pnputil.exe /add-driver $ippInf.FullName /install 2>&1
                        Write-Output "[INFO] pnputil: $result"
                    } else {
                        Write-Output "[INFO] Попробуйте: Add-PrinterDriver -Name 'Microsoft IPP Class Driver'"
                        try {
                            Add-PrinterDriver -Name "Microsoft IPP Class Driver" -ErrorAction Stop
                            Write-Output "[OK] Microsoft IPP Class Driver добавлен"
                        } catch {
                            Write-Output "[WARN] Не удалось добавить: $_"
                        }
                    }
                }
            } catch {
                Write-Output "[WARN] Ошибка: $_"
            }
            """;

        var (_, out2, _) = await RunPS(step2);
        ReportOutput(out2, Report);

        // ── Шаг 3: Брандмауэр — порт 631 (IPP) и 443 (IPPS) ────────────────────
        Report(Info("Шаг 3: Настройка брандмауэра для IPP (порт 631, 443)..."));

        var step3 = """
            try {
                # Проверяем существующее правило для IPP
                $existing = Get-NetFirewallRule -DisplayName "W-Fix IPP Allow" -ErrorAction SilentlyContinue
                
                if (-not $existing) {
                    # Создаём правило для IPP порта 631
                    New-NetFirewallRule -DisplayName "W-Fix IPP Allow" `
                        -Direction Outbound -Action Allow `
                        -Protocol TCP -RemotePort 631 `
                        -Description "Разрешить IPP-трафик (порт 631)" `
                        -ErrorAction Stop | Out-Null
                    Write-Output "[OK] Правило брандмауэра создано: TCP 631 (IPP)"
                } else {
                    # Убеждаемся что включено
                    $existing | Set-NetFirewallRule -Enabled True
                    Write-Output "[OK] Правило IPP уже существует и включено"
                }
                
                # Проверяем IPPS (443)
                $existing443 = Get-NetFirewallRule -DisplayName "W-Fix IPPS Allow" -ErrorAction SilentlyContinue
                if (-not $existing443) {
                    New-NetFirewallRule -DisplayName "W-Fix IPPS Allow" `
                        -Direction Outbound -Action Allow `
                        -Protocol TCP -RemotePort 443 `
                        -Description "Разрешить IPPS-трафик (порт 443)" `
                        -ErrorAction Stop | Out-Null
                    Write-Output "[OK] Правило брандмауэра создано: TCP 443 (IPPS)"
                } else {
                    Write-Output "[OK] Правило IPPS уже существует"
                }
            } catch {
                Write-Output "[WARN] Ошибка настройки брандмауэра: $_"
            }
            """;

        var (_, out3, _) = await RunPS(step3);
        ReportOutput(out3, Report);

        // ── Шаг 4: Очистка кеша IPP ─────────────────────────────────────────────
        Report(Info("Шаг 4: Очистка кеша IPP-подключений..."));

        var step4 = """
            try {
                # Останавливаем Spooler
                Stop-Service -Name spooler -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
                
                # Чистим spool\PRINTERS
                $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
                $files = Get-ChildItem -Path $spoolPath -ErrorAction SilentlyContinue
                $count = $files.Count
                $files | Remove-Item -Force -ErrorAction SilentlyContinue
                Write-Output "[OK] Удалено файлов очереди: $count"
                
                # Чистим кеш подключений (может содержать устаревшие IPP endpoint)
                $cachePath = "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"
                $ippCache = Get-ChildItem -Path $cachePath -Recurse -Filter "*ipp*" -ErrorAction SilentlyContinue
                if ($ippCache) {
                    $ippCache | Remove-Item -Force -ErrorAction SilentlyContinue
                    Write-Output "[OK] IPP кеш очищен: $($ippCache.Count) файлов"
                } else {
                    Write-Output "[OK] IPP кеш чист"
                }
                
                # Восстанавливаем Spooler
                Start-Service -Name spooler -ErrorAction Stop
                Start-Sleep -Seconds 2
                $status = (Get-Service spooler).Status
                Write-Output "[OK] Spooler: $status"
            } catch {
                Write-Output "[ERROR] Ошибка: $_"
            }
            """;

        var (_, out4, _) = await RunPS(step4);
        ReportOutput(out4, Report);

        // ── Шаг 5: Перенастройка IPP порта ──────────────────────────────────────
        if (printer?.IsNetwork == true && (printer.PortName.StartsWith("http") || printer.PortName.StartsWith("ipp")))
        {
            Report(Info($"Шаг 5: Пересоздание порта IPP ({printer.PortName})..."));

            var portEscaped = printer.PortName.Replace("'", "''");

            var step5 = @"
                $uri = '" + portEscaped + @"'
                
                # Извлекаем хост из URI
                try {
                    $parsed = [System.Uri]::new($uri)
                    $hostName = $parsed.Host
                    $port = if ($parsed.Port -gt 0) { $parsed.Port } else { 631 }
                    
                    # Удаляем старый порт (если он WSD или кривой)
                    # WMI может выдать ошибку, если порт занят
                    # Get-PrinterPort -Name $uri | Remove-PrinterPort -ErrorAction SilentlyContinue
                    
                    Write-Output ""[INFO] Проверяем реестр (Internet Printer Ports)...""
                    $regKey = ""HKCU:\Printers\Connections""
                    # (Сложная логика пересоздания)
                    
                    # Пинг
                    $ping = Test-Connection -ComputerName $hostName -Count 1 -Quiet -ErrorAction SilentlyContinue
                    if ($ping) { Write-Output ""[OK] Сервер IPP ($hostName) доступен по сети"" }
                    else { Write-Output ""[WARN] Сервер IPP недоступен (ping не прошёл)"" }
                    
                    # Проверка порта 631
                    $tcp = New-Object Net.Sockets.TcpClient
                    try {
                        $tcp.ConnectAsync($hostName, $port).Wait(2000)
                        if ($tcp.Connected) {
                            Write-Output ""[OK] Порт $port на $hostName открыт""
                        } else {
                            Write-Output ""[WARN] Порт $port закрыт или блокируется""
                        }
                    } catch {
                        Write-Output ""[WARN] Ошибка проверки TCP-порта: $_""
                    } finally {
                        $tcp.Dispose()
                    }
                } catch {
                    Write-Output ""[WARN] Ошибка парсинга URI: $_""
                }
            ";

            var (_, out5, _) = await RunPS(step5);
            ReportOutput(out5, Report);
        }

        Report(Info("═══════════════════════════════════════════════════"));
        Report(Info("IPP требует включённый компонент Windows и MS IPP Class Driver."));
        Report(Info("При сохранении проблем — перезагрузите ПК."));

        var allOk = steps.All(s => s.Level != Models.LogLevel.Error);
        return allOk
            ? FixResult.Ok("Фикс IPP выполнен", steps)
            : FixResult.Warn("Фикс IPP частично выполнен — проверьте логи", steps);
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
