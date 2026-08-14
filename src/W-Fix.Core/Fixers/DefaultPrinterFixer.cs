using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Сброс и переустановка принтера по умолчанию.
/// </summary>
public class DefaultPrinterFixer : FixerBase, ISystemStateChangingFixer
{
    private readonly InteractiveUserPowerShellService _interactiveUserPowerShell = new();

    public override string Name => "Сброс принтера по умолчанию";
    public override string Description =>
        "Очищает некорректную запись PrinterDevice в реестре и переустанавливает принтер по умолчанию. " +
        "Решает проблему «принтер по умолчанию не сохраняется» или выбирается неверная бумага.";
    public override string[] TargetErrorCodes => ["default", "device", "HKCU_printer"];

    public SystemStateBackupPlan CreateBackupPlan(PrinterInfo? printer) => new()
    {
        RegistryValues =
        [
            new(@"HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows", "LegacyDefaultPrinterMode"),
            new(@"HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows", "UserSelectedDefault"),
            new(@"HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device")
        ]
    };

    public override async Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        var newDefault = printer?.Name ?? "";
        Report(Info(string.IsNullOrEmpty(newDefault)
            ? "Сброс записей принтера по умолчанию для целевого пользователя..."
            : $"Установка принтера по умолчанию: '{newDefault}'"));

        var script = @"
            $ErrorActionPreference = 'Stop'
            $newDefault = '" + newDefault.Replace("'", "''") + @"'

            # Отключить ""автоматический выбор принтера по умолчанию"" (Windows 10+)
            $regPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows'
            $legacyVal = (Get-ItemProperty -Path $regPath -Name LegacyDefaultPrinterMode -ErrorAction SilentlyContinue).LegacyDefaultPrinterMode
            if ($legacyVal -ne 1) {
                Set-ItemProperty -Path $regPath -Name LegacyDefaultPrinterMode -Value 1 -Type DWord -Force
                Write-Output ""[OK] Отключён автовыбор принтера по умолчанию (LegacyDefaultPrinterMode=1)""
            } else {
                Write-Output ""[INFO] LegacyDefaultPrinterMode уже = 1""
            }

            if ($newDefault -eq '') {
                foreach ($valueName in @('Device', 'UserSelectedDefault')) {
                    $key = Get-Item -LiteralPath $regPath -ErrorAction Stop
                    if ($key.GetValueNames() -contains $valueName) {
                        Remove-ItemProperty -Path $regPath -Name $valueName -ErrorAction Stop
                    }
                }
                $remainingNames = (Get-Item -LiteralPath $regPath -ErrorAction Stop).GetValueNames()
                if ($remainingNames -contains 'Device' -or $remainingNames -contains 'UserSelectedDefault') {
                    Write-Output ""[ERROR] Проверка сброса Device/UserSelectedDefault не пройдена""
                    exit 1
                }
                Write-Output ""[OK] Записи Device и UserSelectedDefault сброшены""
            }

            # Установить конкретный принтер как default
            if ($newDefault -ne '') {
                try {
                    (New-Object -ComObject WScript.Network).SetDefaultPrinter($newDefault)
                    Write-Output ""[OK] Принтер по умолчанию установлен: $newDefault""
                } catch {
                    Write-Output ""[WARN] WScript.Network не сработал, пробуем через rundll32...""
                    & rundll32 printui.dll,PrintUIEntry /y /n $newDefault 2>&1 | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        Write-Output ""[ERROR] PrintUI завершился с кодом $LASTEXITCODE""
                        exit 1
                    }
                    Write-Output ""[OK] PrintUI выполнен для $newDefault""
                }
            }

            # Проверяем результат
            $check = (Get-CimInstance -Class Win32_Printer | Where-Object { $_.Default -eq $true }).Name
            if ($newDefault -ne '' -and $check -ne $newDefault) {
                Write-Output ""[ERROR] Проверка не пройдена: текущий принтер '$check', ожидался '$newDefault'""
                exit 1
            }
            Write-Output ""[OK] Текущий принтер по умолчанию: $check""
        ";

        var execution = remoteMachine == null
            ? await PowerShellEngine.RunExternalAsync(script, ct)
            : await _interactiveUserPowerShell.RunRemoteAsync(remoteMachine, script, ct);

        foreach (var line in execution.Output)
        {
            var level = line.StartsWith("[OK]") ? Models.LogLevel.Success
                      : line.StartsWith("[WARN]") ? Models.LogLevel.Warning
                      : line.StartsWith("[ERROR]") ? Models.LogLevel.Error
                      : Models.LogLevel.Info;
            Report(new LogEntry(level, line));
        }

        if (!execution.Success)
            return FixResult.Fail($"Не удалось изменить принтер по умолчанию: {execution.Error}", steps);

        var summary = string.IsNullOrEmpty(newDefault)
                ? "Настройки принтера по умолчанию обновлены"
                : $"Принтер '{newDefault}' назначен по умолчанию";
        return execution.Output.Any(line => line.StartsWith("[WARN]", StringComparison.OrdinalIgnoreCase))
            ? FixResult.Warn(summary + " с предупреждениями", steps)
            : FixResult.Ok(summary, steps);
    }
}
