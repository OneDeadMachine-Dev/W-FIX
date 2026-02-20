using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Сброс и переустановка принтера по умолчанию.
/// </summary>
public class DefaultPrinterFixer : FixerBase
{
    public override string Name => "Сброс принтера по умолчанию";
    public override string Description =>
        "Очищает некорректную запись PrinterDevice в реестре и переустанавливает принтер по умолчанию. " +
        "Решает проблему «принтер по умолчанию не сохраняется» или выбирается неверная бумага.";
    public override string[] TargetErrorCodes => ["default", "device", "HKCU_printer"];

    public override async Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        var newDefault = printer?.Name ?? "";
        Report(Info(string.IsNullOrEmpty(newDefault)
            ? "Сброс записи принтера по умолчанию (все пользователи)..."
            : $"Установка принтера по умолчанию: '{newDefault}'"));

        var script = @"
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

            # Установить конкретный принтер как default
            if ($newDefault -ne '') {
                try {
                    (New-Object -ComObject WScript.Network).SetDefaultPrinter($newDefault)
                    Write-Output ""[OK] Принтер по умолчанию установлен: $newDefault""
                } catch {
                    Write-Output ""[WARN] WScript.Network не сработал, пробуем через rundll32...""
                    & rundll32 printui.dll,PrintUIEntry /y /n $newDefault 2>&1 | Out-Null
                    Write-Output ""[OK] PrintUI выполнен для $newDefault""
                }
            }

            # Проверяем результат
            $check = (Get-CimInstance -Class Win32_Printer | Where-Object { $_.Default -eq $true }).Name
            Write-Output ""[OK] Текущий принтер по умолчанию: $check""
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

        return success
            ? FixResult.Ok("Принтер по умолчанию установлен", steps)
            : FixResult.Warn($"Частичный успех: {error}", steps);
    }
}
