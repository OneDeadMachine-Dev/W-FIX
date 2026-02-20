using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Фикс ошибки 0x0000011b (ERROR_INVALID_PRINTER_NAME / RPC Authentication).
/// Появилась после Windows Update KB5005565 (Windows 10/11 2021+).
/// Суть: Windows ввёл строгую аутентификацию RPC при подключении к сетевым принтерам.
/// Решение: установить RpcAuthnLevelPrivacyEnabled = 0 на сервере печати.
/// </summary>
public class Error11bFixer : FixerBase
{
    private const string RegPath = @"HKLM:\System\CurrentControlSet\Control\Print";
    private const string RegValue = "RpcAuthnLevelPrivacyEnabled";

    public override string Name => "Ошибка 0x0000011b (RPC Auth)";
    public override string Description =>
        "Устанавливает RpcAuthnLevelPrivacyEnabled = 0 в реестре сервера печати и перезапускает Spooler. " +
        "Решает ошибку подключения к сетевому принтеру после обновлений Windows (KB5005565+).";
    public override string[] TargetErrorCodes => ["0x0000011b", "11b", "ERROR_INVALID_PRINTER_NAME", "0x0000011B"];

    public override async Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        Report(Info($"Применение патча реестра {RegValue} на {remoteMachine ?? "локальной машине"}..."));

        // Используем @"" verbatim strings — PowerShell {} не конфликтуют с C# интерполяцией
        var script = @"
            $regPath = '" + RegPath + @"'
            $regValue = '" + RegValue + @"'

            # Проверяем текущее значение
            $current = (Get-ItemProperty -Path $regPath -Name $regValue -ErrorAction SilentlyContinue).$regValue
            Write-Output ""[INFO] Текущее значение: $current""

            if ($current -eq 0) {
                Write-Output ""[OK] Патч уже применён (значение = 0). Дополнительных действий не требуется.""
            } else {
                # Применяем патч
                Set-ItemProperty -Path $regPath -Name $regValue -Value 0 -Type DWord -Force
                $verify = (Get-ItemProperty -Path $regPath -Name $regValue).$regValue
                if ($verify -eq 0) {
                    Write-Output ""[OK] Реестр обновлён: $regValue = 0""
                } else {
                    Write-Output ""[ERROR] Не удалось применить патч реестра!""
                    exit 1
                }
            }

            # Перезапуск Spooler
            Write-Output ""[INFO] Перезапускаем Print Spooler...""
            Stop-Service -Name spooler -Force -ErrorAction SilentlyContinue
            Start-Service -Name spooler -ErrorAction Stop
            $status = (Get-Service spooler).Status
            Write-Output ""[OK] Spooler статус: $status""
            Write-Output ""[OK] Патч 0x0000011b применён успешно""
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

        if (success)
        {
            Report(Ok("Рекомендация: на клиентских ПК можно также применить этот патч для полного устранения."));
            return FixResult.Ok("Патч 0x0000011b успешно применён", steps);
        }
        return FixResult.Fail($"Ошибка патча 0x0000011b: {error}", steps);
    }
}
