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
public class Error709Fixer : FixerBase, ISystemStateChangingFixer
{
    private readonly InteractiveUserPowerShellService _interactiveUserPowerShell = new();

    public override string Name => "Ошибка 0x00000709 (Принтер по умолчанию)";
    public override string Description =>
        "Исправляет ошибку «Невозможно завершить операцию (0x00000709)» при назначении принтера по умолчанию. " +
        "Сбрасывает политику автоуправления, чинит права реестра и переназначает принтер.";
    public override string[] TargetErrorCodes => ["0x00000709", "709", "default_printer_failed"];

    public SystemStateBackupPlan CreateBackupPlan(PrinterInfo? printer)
    {
        const string windowsKey = @"HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows";
        return new SystemStateBackupPlan
        {
            RegistryValues =
            [
                new(windowsKey, "LegacyDefaultPrinterMode"),
                new(windowsKey, "UserSelectedDefault"),
                new(windowsKey, "Device")
            ],
            RegistryAcls = [new RegistryAclBackupTarget(windowsKey)]
        };
    }

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
            $ErrorActionPreference = 'Stop'
            $printerName = '" + printer.Name.Replace("'", "''") + @"'
            $remoteInteractive = " + (remoteMachine == null ? "$false" : "$true") + @"

            # ── Шаг 1: Отключить автоуправление принтером по умолчанию ──
            Write-Output ""[INFO] Шаг 1: Отключение автоуправления принтером по умолчанию...""
            try {
                $regPath = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows'
                
                # Отключаем LegacyDefaultPrinterMode = 1 (Windows не управляет сам)
                $legacyPath = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows'
                Set-ItemProperty -Path $legacyPath -Name 'LegacyDefaultPrinterMode' -Value 1 -Type DWord -Force -ErrorAction Stop
                Write-Output ""[OK] LegacyDefaultPrinterMode = 1 (ручной режим)""
            } catch {
                Write-Output ""[WARN] Не удалось изменить LegacyDefaultPrinterMode: $_""
            }

            # ── Шаг 2: Починить права на ключ реестра ──
            Write-Output ""[INFO] Шаг 2: Проверка прав на ключ реестра...""
            if ($remoteInteractive) {
                Write-Output ""[INFO] ACL подготовлен административным контекстом перед запуском пользовательского сценария""
            } else { try {
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
            } }

            # ── Шаг 3: Очистить старое значение Device ──
            Write-Output ""[INFO] Шаг 3: Очистка записи Device...""
            try {
                $regPath = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows'
                
                # Удаляем UserSelectDefault (если есть — мешает)
                if ((Get-Item -LiteralPath $regPath).GetValueNames() -contains 'UserSelectedDefault') {
                    Remove-ItemProperty -Path $regPath -Name 'UserSelectedDefault' -ErrorAction Stop
                    Write-Output ""[OK] UserSelectedDefault удалён""
                } else {
                    Write-Output ""[INFO] UserSelectedDefault отсутствует""
                }
                
                # Пишем корректное значение Device
                $wmiPrinter = Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -eq $printerName } | Select-Object -First 1
                if ($wmiPrinter) {
                    $port = $wmiPrinter.PortName
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
                    $cimPrinter = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop |
                        Where-Object { $_.Name -eq $printerName } | Select-Object -First 1
                    if (-not $cimPrinter) {
                        throw ""Принтер '$printerName' не найден в CIM""
                    }
                    Invoke-CimMethod -InputObject $cimPrinter -MethodName SetDefaultPrinter | Out-Null
                    Write-Output ""[OK] Принтер '$printerName' установлен по умолчанию (CIM)""
                } catch {
                    Write-Output ""[ERROR] Не удалось назначить принтер: $_""
                    exit 1
                }
            }

            # ── Шаг 5: Перезапуск Spooler ──
            if ($remoteInteractive) {
                Write-Output ""[INFO] Шаг 5: Spooler будет перезапущен административным контекстом""
            } else {
                Write-Output ""[INFO] Шаг 5: Перезапуск Print Spooler...""
                Restart-Service -Name spooler -Force -ErrorAction Stop
                Start-Sleep -Seconds 2
                $status = (Get-Service spooler -ErrorAction Stop).Status
                if ($status -ne 'Running') {
                    Write-Output ""[ERROR] Spooler не перешёл в состояние Running""
                    exit 1
                }
                Write-Output ""[OK] Spooler перезапущен: $status""
            }

            # Проверка
            $default = Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue |
                Where-Object { $_.Default -eq $true } | Select-Object -First 1
            if (-not $default -or $default.Name -ne $printerName) {
                Write-Output ""[ERROR] Проверка не пройдена: текущий принтер '$($default.Name)', ожидался '$printerName'""
                exit 1
            }
            Write-Output ""[OK] Принтер по умолчанию подтверждён: $($default.Name)""
        ";

        if (remoteMachine == null)
        {
            var localExecution = await PowerShellEngine.RunExternalAsync(script, ct);
            ReportOutput(localExecution.Output, Report);
            if (!localExecution.Success)
                return FixResult.Fail($"Ошибка: {localExecution.Error}", steps);

            return HasWarnings(localExecution.Output)
                ? FixResult.Warn($"Принтер '{printer.Name}' назначен с предупреждениями", steps)
                : FixResult.Ok($"Принтер '{printer.Name}' назначен по умолчанию", steps);
        }

        var aclExecution = await PrepareRemoteUserRegistryAclAsync(remoteMachine, ct);
        ReportOutput(aclExecution.Output, Report);

        var userExecution = await _interactiveUserPowerShell.RunRemoteAsync(remoteMachine, script, ct);
        ReportOutput(userExecution.Output, Report);
        if (!userExecution.Success)
            return FixResult.Fail($"Не удалось назначить принтер в сеансе пользователя: {userExecution.Error}", steps);

        var spoolerExecution = await RestartRemoteSpoolerAsync(remoteMachine, ct);
        ReportOutput(spoolerExecution.Output, Report);

        if (!aclExecution.Success || !spoolerExecution.Success || HasWarnings(userExecution.Output))
        {
            var warnings = new List<string>();
            if (!aclExecution.Success) warnings.Add($"ACL: {aclExecution.Error}");
            if (!spoolerExecution.Success) warnings.Add($"Spooler: {spoolerExecution.Error}");
            if (HasWarnings(userExecution.Output)) warnings.Add("пользовательский сценарий завершён с предупреждениями");
            return FixResult.Warn(
                $"Принтер '{printer.Name}' назначен, но есть предупреждения: {string.Join("; ", warnings)}",
                steps);
        }

        return FixResult.Ok($"Принтер '{printer.Name}' назначен по умолчанию", steps);
    }

    private static bool HasWarnings(IReadOnlyList<string> output) =>
        output.Any(line => line.StartsWith("[WARN]", StringComparison.OrdinalIgnoreCase));

    private static async Task<PowerShellExecutionResult> PrepareRemoteUserRegistryAclAsync(
        string remoteMachine,
        CancellationToken ct)
    {
        const string script = """
            $ErrorActionPreference = 'Stop'
            try {
                $interactiveUser = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).UserName
                if ([string]::IsNullOrWhiteSpace($interactiveUser)) {
                    throw 'На целевом компьютере нет вошедшего интерактивного пользователя.'
                }
                $sid = ([System.Security.Principal.NTAccount]$interactiveUser).Translate(
                    [System.Security.Principal.SecurityIdentifier])
                $path = "Registry::HKEY_USERS\$sid\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows"
                if (-not (Test-Path -LiteralPath $path)) {
                    throw "Пользовательский раздел реестра не загружен: $path"
                }

                $acl = Get-Acl -LiteralPath $path -ErrorAction Stop
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule(
                    $sid, 'FullControl', 'Allow')
                $acl.SetAccessRule($rule)
                Set-Acl -LiteralPath $path -AclObject $acl -ErrorAction Stop
                Write-Output "[OK] ACL HKCU восстановлен для $interactiveUser"
            } catch {
                Write-Output "[ERROR] Не удалось подготовить ACL HKCU: $_"
                exit 1
            }
            """;

        using var engine = new PowerShellEngine(remoteMachine);
        return await engine.RunAsync(script, ct: ct);
    }

    private static async Task<PowerShellExecutionResult> RestartRemoteSpoolerAsync(
        string remoteMachine,
        CancellationToken ct)
    {
        const string script = """
            $ErrorActionPreference = 'Stop'
            try {
                Restart-Service -Name spooler -Force -ErrorAction Stop
                $service = Get-Service -Name spooler -ErrorAction Stop
                $service.WaitForStatus('Running', [TimeSpan]::FromSeconds(20))
                Write-Output "[OK] Spooler удалённой машины перезапущен: $($service.Status)"
            } catch {
                Write-Output "[ERROR] Не удалось перезапустить Spooler: $_"
                exit 1
            }
            """;

        using var engine = new PowerShellEngine(remoteMachine);
        return await engine.RunAsync(script, ct: ct, timeout: TimeSpan.FromSeconds(30));
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
