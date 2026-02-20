using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Фикс ошибки 0x00000709 — «Невозможно завершить операцию» при установке принтера по умолчанию.
/// Причины: Windows сам управляет принтером по умолчанию, повреждены ключи реестра,
/// нет прав на запись в HKCU\...\Windows, или конфликт имён.
///
/// Шаги:
///   1. Отключение "Let Windows manage my default printer"
///   2. Выдача Full Control на ключ реестра HKCU\..\Windows
///   3. Очистка значения Device и UserSelectDefault
///   4. Установка принтера по умолчанию через WScript.Network
///   5. Перезапуск Spooler
/// </summary>
public class Error709Fixer : FixerBase
{
    public override string Name => "Ошибка 0x00000709 (Принтер по умолчанию)";
    public override string Description =>
        "Исправляет ошибку «Невозможно завершить операцию (0x00000709)» при назначении принтера по умолчанию. " +
        "Сбрасывает политику автоуправления, чинит права реестра и переназначает принтер.";
    public override string[] TargetErrorCodes => ["0x00000709", "709", "default_printer_failed"];

    public override async Task<FixResult> ApplyAsync(
        PrinterInfo? printer, string? remoteMachine,
        IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        if (printer == null)
        {
            Report(Warn("Выберите принтер для назначения по умолчанию."));
            return FixResult.Warn("Принтер не выбран", steps);
        }

        Report(Info($"Исправление ошибки 0x00000709 для: {printer.Name}"));

        var script = @"
            $printerName = '" + printer.Name.Replace("'", "''") + @"'

            # ── Шаг 1: Отключить автоуправление принтером по умолчанию ──
            Write-Output ""[INFO] Шаг 1: Отключение автоуправления принтером по умолчанию...""
            try {
                $regPath = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows'
                
                # Отключаем LegacyDefaultPrinterMode = 1 (Windows не управляет сам)
                $legacyPath = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows'
                Set-ItemProperty -Path $legacyPath -Name 'LegacyDefaultPrinterMode' -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
                Write-Output ""[OK] LegacyDefaultPrinterMode = 1 (ручной режим)""
            } catch {
                Write-Output ""[WARN] Не удалось изменить LegacyDefaultPrinterMode: $_""
            }

            # ── Шаг 2: Починить права на ключ реестра ──
            Write-Output ""[INFO] Шаг 2: Проверка прав на ключ реестра...""
            try {
                $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey(
                    'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows', $true)
                if ($key) {
                    $acl = $key.GetAccessControl()
                    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                    $rule = New-Object System.Security.AccessControl.RegistryAccessRule(
                        $user, 'FullControl', 'Allow')
                    $acl.SetAccessRule($rule)
                    $key.SetAccessControl($acl)
                    $key.Close()
                    Write-Output ""[OK] Права FullControl установлены для $user""
                } else {
                    Write-Output ""[WARN] Не удалось открыть ключ реестра""
                }
            } catch {
                Write-Output ""[WARN] Ошибка установки прав: $_""
            }

            # ── Шаг 3: Очистить старое значение Device ──
            Write-Output ""[INFO] Шаг 3: Очистка записи Device...""
            try {
                $regPath = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows'
                
                # Удаляем UserSelectDefault (если есть — мешает)
                Remove-ItemProperty -Path $regPath -Name 'UserSelectedDefault' -ErrorAction SilentlyContinue
                Write-Output ""[OK] UserSelectedDefault удалён""
                
                # Пишем корректное значение Device
                $wmiPrinter = Get-CimInstance -ClassName Win32_Printer -Filter ""Name='$printerName'"" -ErrorAction SilentlyContinue
                if ($wmiPrinter) {
                    $port = $wmiPrinter.PortName
                    $driver = $wmiPrinter.DriverName
                    $deviceValue = ""$printerName,winspool,$port""
                    Set-ItemProperty -Path $regPath -Name 'Device' -Value $deviceValue -Type String -Force
                    Write-Output ""[OK] Device = $deviceValue""
                } else {
                    Write-Output ""[WARN] Принтер '$printerName' не найден в WMI""
                }
            } catch {
                Write-Output ""[WARN] Ошибка очистки Device: $_""
            }

            # ── Шаг 4: Установить принтер по умолчанию ──
            Write-Output ""[INFO] Шаг 4: Назначение принтера по умолчанию...""
            try {
                # Метод 1: через WScript.Network
                $net = New-Object -ComObject WScript.Network
                $net.SetDefaultPrinter($printerName)
                Write-Output ""[OK] Принтер '$printerName' установлен по умолчанию (WScript.Network)""
            } catch {
                Write-Output ""[WARN] WScript.Network не сработал: $_""
                try {
                    # Метод 2: через CIM
                    $cimPrinter = Get-CimInstance -ClassName Win32_Printer -Filter ""Name='$printerName'""
                    Invoke-CimMethod -InputObject $cimPrinter -MethodName SetDefaultPrinter | Out-Null
                    Write-Output ""[OK] Принтер '$printerName' установлен по умолчанию (CIM)""
                } catch {
                    Write-Output ""[ERROR] Не удалось назначить принтер: $_""
                    exit 1
                }
            }

            # ── Шаг 5: Перезапуск Spooler ──
            Write-Output ""[INFO] Шаг 5: Перезапуск Print Spooler...""
            Restart-Service -Name spooler -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            $status = (Get-Service spooler).Status
            Write-Output ""[OK] Spooler: $status""

            # Проверка
            $default = Get-CimInstance -ClassName Win32_Printer -Filter ""Default=True"" -ErrorAction SilentlyContinue
            if ($default) {
                Write-Output ""[OK] Принтер по умолчанию: $($default.Name)""
            }
        ";

        // Get-CimInstance требует CimCmdlets (Desktop-only) — используем внешний powershell.exe
        var (success, output, error) = remoteMachine == null
            ? await PowerShellEngine.RunExternalAsync(script, ct)
            : await new PowerShellEngine(remoteMachine).RunAsync(script, ct: ct);
        ReportOutput(output, Report);

        return success
            ? FixResult.Ok($"Принтер '{printer.Name}' назначен по умолчанию", steps)
            : FixResult.Fail($"Ошибка: {error}", steps);
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
