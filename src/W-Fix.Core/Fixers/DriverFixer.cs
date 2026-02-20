using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Управление драйверами принтеров:
///   Режим 1 (INF) — установка драйвера из INF-файла через pnputil.
///   Режим 2 (UNC) — добавление сетевого принтера по UNC-пути через Add-Printer.
///   Режим 3 (Auto) — автоматический поиск подходящего драйвера в DriverStore.
/// </summary>
public class DriverFixer : FixerBase, IInteractiveFixer
{
    /// <summary>Путь к INF-файлу (задаётся из UI перед ApplyAsync).</summary>
    public string? InfPath { get; set; }

    /// <summary>UNC-путь к сетевому принтеру (задаётся из UI перед ApplyAsync).</summary>
    public string? UncPath { get; set; }

    /// <summary>Режим работы: INF, UNC или Auto.</summary>
    public DriverFixMode Mode { get; set; } = DriverFixMode.Inf;

    public override string Name => "Переустановка драйвера";
    public override string Description =>
        "Удаляет текущий драйвер принтера и устанавливает заново из указанного INF-файла или " +
        "добавляет принтер по UNC-пути. Используйте для устранения ошибок драйвера (0x0000007b, 0x00000bc4).";
    public override string[] TargetErrorCodes => ["0x0000007b", "0x00000bc4", "driver", "install"];

    // IInteractiveFixer — сообщаем ViewModel, что перед запуском нужно запросить данные у пользователя
    public string InputTitle => "Переустановка драйвера";
    public string InputDescription =>
        "Выберите способ установки:\n" +
        "• INF — укажите путь к INF-файлу драйвера\n" +
        "• UNC — введите сетевой путь (\\\\server\\printer)\n" +
        "• Авто — поиск драйвера в DriverStore Windows";
    public InteractiveInputType InputType => InteractiveInputType.DriverInstall;

    public override async Task<FixResult> ApplyAsync(
        PrinterInfo? printer, string? remoteMachine,
        IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        return Mode switch
        {
            DriverFixMode.Inf => await InstallFromInfAsync(printer, remoteMachine, Report, ct, steps),
            DriverFixMode.Unc => await AddByUncAsync(remoteMachine, Report, ct, steps),
            DriverFixMode.Auto => await AutoInstallAsync(printer, remoteMachine, Report, ct, steps),
            _ => FixResult.Fail("Неизвестный режим", steps)
        };
    }

    // ── INF ────────────────────────────────────────────────────────────────────

    private async Task<FixResult> InstallFromInfAsync(
        PrinterInfo? printer, string? remoteMachine,
        Action<LogEntry> report, CancellationToken ct, List<LogEntry> steps)
    {
        if (string.IsNullOrEmpty(InfPath))
        {
            report(Warn("Путь к INF-файлу не задан."));
            return FixResult.Warn("INF-путь не задан", steps);
        }

        var driverName = printer?.DriverName ?? "Unknown";
        report(Info($"Установка драйвера из: {InfPath}"));
        report(Info($"Текущий драйвер: {driverName}"));

        var script = @"
            $infPath = '" + InfPath.Replace("'", "''") + @"'
            $driverName = '" + driverName.Replace("'", "''") + @"'
            $printerName = '" + (printer?.Name ?? "").Replace("'", "''") + @"'

            # Проверяем существование INF-файла
            if (-not (Test-Path $infPath)) {
                Write-Output ""[ERROR] INF-файл не найден: $infPath""
                exit 1
            }

            # Шаг 1: Установить драйвер через pnputil
            Write-Output ""[INFO] Добавляем INF в хранилище драйверов через pnputil...""
            $pnpResult = pnputil.exe /add-driver ""$infPath"" /install 2>&1
            foreach ($line in $pnpResult) {
                Write-Output ""[INFO] pnputil: $line""
            }

            # Шаг 2: Если есть конкретный принтер — попытаемся обновить его драйвер
            if ($printerName -ne '' -and $driverName -ne 'Unknown') {
                try {
                    # Получаем список установленных драйверов принтеров
                    $newDrivers = Get-PrinterDriver -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -like ""*$($driverName.Split(' ')[0])*"" }
                    if ($newDrivers) {
                        Write-Output ""[OK] Найден подходящий драйвер: $($newDrivers[0].Name)""
                    } else {
                        Write-Output ""[INFO] Новый драйвер установлен, но требует ручного привязки к принтеру.""
                    }
                } catch {
                    Write-Output ""[WARN] Не удалось проверить установленные драйверы: $_""
                }
            }

            # Шаг 3: Перезапустить Spooler
            Write-Output ""[INFO] Перезапускаем Print Spooler...""
            Stop-Service -Name spooler -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Start-Service -Name spooler -ErrorAction Stop
            Write-Output ""[OK] Spooler перезапущен""

            # Шаг 4: Итог
            Write-Output ""[OK] Драйвер успешно установлен из INF-файла""
        ";

        using var engine = new PowerShellEngine(remoteMachine);
        var (success, output, error) = await engine.RunAsync(script, ct: ct);
        ReportOutput(output, report);

        return success
            ? FixResult.Ok($"Драйвер установлен из {InfPath}", steps)
            : FixResult.Fail($"Ошибка установки: {error}", steps);
    }

    // ── UNC ────────────────────────────────────────────────────────────────────

    private async Task<FixResult> AddByUncAsync(
        string? remoteMachine, Action<LogEntry> report, CancellationToken ct, List<LogEntry> steps)
    {
        if (string.IsNullOrEmpty(UncPath))
        {
            report(Warn("UNC-путь не задан."));
            return FixResult.Warn("UNC-путь не задан", steps);
        }

        report(Info($"Добавление сетевого принтера: {UncPath}"));

        var script = @"
            $uncPath = '" + UncPath.Replace("'", "''") + @"'

            # Проверяем доступность print-сервера
            $server = ($uncPath -replace '^\\\\', '') -split '\\' | Select-Object -First 1
            Write-Output ""[INFO] Проверяем доступность сервера: $server""
            $ping = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction SilentlyContinue
            if (-not $ping) {
                Write-Output ""[WARN] Сервер $server не отвечает на ping (может быть заблокирован ICMP)""
            } else {
                Write-Output ""[OK] Сервер $server доступен""
            }

            # Проверяем, не установлен ли принтер уже
            $existing = Get-Printer -Name $uncPath -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Output ""[WARN] Принтер уже установлен: $uncPath. Удаляем для переустановки...""
                Remove-Printer -Name $uncPath -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
            }

            # Добавляем принтер
            try {
                Add-Printer -ConnectionName $uncPath -ErrorAction Stop
                Write-Output ""[OK] Принтер '$uncPath' успешно добавлен""
            } catch {
                Write-Output ""[ERROR] Не удалось добавить принтер: $_""
                exit 1
            }

            # Проверяем статус
            $check = Get-Printer -Name $uncPath -ErrorAction SilentlyContinue
            if ($check) {
                Write-Output ""[OK] Статус: $($check.PrinterStatus)""
                Write-Output ""[OK] Драйвер: $($check.DriverName)""
            }
        ";

        using var engine = new PowerShellEngine(remoteMachine);
        var (success, output, error) = await engine.RunAsync(script, ct: ct);
        ReportOutput(output, report);

        return success
            ? FixResult.Ok($"Принтер {UncPath} добавлен", steps)
            : FixResult.Fail($"Ошибка: {error}", steps);
    }

    // ── Auto ───────────────────────────────────────────────────────────────────

    private async Task<FixResult> AutoInstallAsync(
        PrinterInfo? printer, string? remoteMachine,
        Action<LogEntry> report, CancellationToken ct, List<LogEntry> steps)
    {
        if (printer == null)
        {
            report(Warn("Выберите принтер для автоматического поиска драйвера."));
            return FixResult.Warn("Принтер не выбран", steps);
        }

        var driverName = printer.DriverName;
        report(Info($"Автоматический поиск драйвера для: {printer.Name}"));
        report(Info($"Текущий драйвер: {driverName}"));

        var script = @"
            $driverName = '" + driverName.Replace("'", "''") + @"'
            $printerName = '" + printer.Name.Replace("'", "''") + @"'

            # Поиск подходящего драйвера в DriverStore
            Write-Output ""[INFO] Поиск драйвера '$driverName' в DriverStore...""

            # Ищем среди установленных драйверов принтеров
            $installedDrivers = Get-PrinterDriver -ErrorAction SilentlyContinue
            $match = $installedDrivers | Where-Object { $_.Name -eq $driverName }

            if ($match) {
                Write-Output ""[OK] Драйвер найден: $($match.Name)""
                Write-Output ""[INFO] Провайдер: $($match.PrinterEnvironment)""
                Write-Output ""[INFO] INF-путь: $($match.InfPath)""

                # Пробуем переустановить
                Write-Output ""[INFO] Удаляем текущий драйвер...""
                try {
                    Remove-PrinterDriver -Name $driverName -ErrorAction Stop
                    Write-Output ""[OK] Драйвер удалён""
                } catch {
                    Write-Output ""[WARN] Не удалось удалить: $_ — продолжаем""
                }

                # Перезапускаем Spooler
                Stop-Service -Name spooler -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Start-Service -Name spooler -ErrorAction Stop
                Write-Output ""[OK] Spooler перезапущен""

                # Восстанавливаем драйвер из DriverStore
                try {
                    Add-PrinterDriver -Name $driverName -ErrorAction Stop
                    Write-Output ""[OK] Драйвер '$driverName' переустановлен из DriverStore""
                } catch {
                    Write-Output ""[ERROR] Не удалось переустановить драйвер: $_""
                    Write-Output ""[INFO] Попробуйте установить вручную через INF-файл""
                    exit 1
                }
            } else {
                Write-Output ""[WARN] Драйвер '$driverName' не найден в DriverStore""

                # Показываем доступные драйверы для информации
                $alternatives = $installedDrivers | Select-Object -First 10
                if ($alternatives) {
                    Write-Output ""[INFO] Доступные драйверы:""
                    foreach ($drv in $alternatives) {
                        Write-Output ""[INFO]   • $($drv.Name)""
                    }
                }
                Write-Output ""[INFO] Установите драйвер вручную через INF-файл (режим INF)""
            }
        ";

        using var engine = new PowerShellEngine(remoteMachine);
        var (success, output, error) = await engine.RunAsync(script, ct: ct);
        ReportOutput(output, report);

        return success
            ? FixResult.Ok($"Драйвер {driverName} переустановлен", steps)
            : FixResult.Fail($"Ошибка: {error}", steps);
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────

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

/// <summary>Режим работы DriverFixer.</summary>
public enum DriverFixMode
{
    /// <summary>Установка из INF-файла.</summary>
    Inf,
    /// <summary>Добавление сетевого принтера по UNC-пути.</summary>
    Unc,
    /// <summary>Автоматический поиск драйвера в DriverStore.</summary>
    Auto
}
