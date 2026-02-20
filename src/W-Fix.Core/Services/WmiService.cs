using System.Management;
using WFix.Core.Models;

namespace WFix.Core.Services;

/// <summary>
/// Сервис для получения информации о принтерах и очередях через WMI.
/// Для локальной машины: прямой WMI без подключения scope (самый надёжный вариант).
/// Для удалённой машины: WMI over DCOM.
/// </summary>
public class WmiService
{
    // ── Публичный API ─────────────────────────────────────────────────────────

    /// <summary>
    /// Список принтеров: локально или на удалённой машине.
    /// </summary>
    public IReadOnlyList<PrinterInfo> GetPrinters(string? machineName = null,
        string? username = null, string? password = null)
    {
        return string.IsNullOrEmpty(machineName)
            ? GetLocalPrinters()
            : GetRemotePrinters(machineName, username, password);
    }

    /// <summary>Список заданий очереди принтера.</summary>
    public IReadOnlyList<PrintJobInfo> GetPrintJobs(string printerName,
        string? machineName = null)
    {
        var result = new List<PrintJobInfo>();
        try
        {
            // Win32_PrintJob.Name = "PrinterName,JobId"
            var safeN = printerName.Replace("'", "\\'");
            var query = $"SELECT * FROM Win32_PrintJob WHERE Name LIKE '{safeN}%'";

            using var searcher = string.IsNullOrEmpty(machineName)
                ? new ManagementObjectSearcher(query)
                : new ManagementObjectSearcher(
                    BuildRemoteScope(machineName),
                    new ObjectQuery(query));

            foreach (ManagementObject obj in searcher.Get())
            {
                result.Add(new PrintJobInfo
                {
                    JobId      = safe32(obj["JobId"]),
                    Document   = str(obj["Document"]),
                    Owner      = str(obj["Owner"]),
                    Status     = str(obj["Status"]),
                    TotalPages = safe32(obj["TotalPages"]),
                    Size       = Convert.ToInt64(obj["Size"] ?? 0L),
                });
            }
        }
        catch (Exception ex)
        {
            Serilog.Log.Error(ex, "WMI GetPrintJobs [{Printer}]", printerName);
        }
        return result;
    }

    /// <summary>Список установленных драйверов принтеров.</summary>
    public IReadOnlyList<PrinterDriverInfo> GetInstalledDrivers(string? machineName = null)
    {
        var result = new List<PrinterDriverInfo>();
        try
        {
            using var searcher = string.IsNullOrEmpty(machineName)
                ? new ManagementObjectSearcher("SELECT * FROM Win32_PrinterDriver")
                : new ManagementObjectSearcher(
                    BuildRemoteScope(machineName),
                    new ObjectQuery("SELECT * FROM Win32_PrinterDriver"));

            foreach (ManagementObject obj in searcher.Get())
            {
                result.Add(new PrinterDriverInfo
                {
                    Name    = str(obj["Name"]),
                    Version = str(obj["DriverVersion"]),
                    InfName = str(obj["InfName"]),
                });
            }
        }
        catch (Exception ex)
        {
            Serilog.Log.Error(ex, "WMI GetInstalledDrivers [{Machine}]",
                machineName ?? "local");
        }
        return result;
    }

    // ── Внутренние методы ────────────────────────────────────────────────────

    /// <summary>
    /// Локальные принтеры с тройным fallback:
    ///   1. Get-CimInstance (PowerShell, не требует модулей — самый портативный)
    ///   2. WMI ManagementObjectSearcher (нативный COM — может не работать в single-file)
    ///   3. Get-Printer PowerShell cmdlet (требует модуль PrintManagement)
    /// </summary>
    private static IReadOnlyList<PrinterInfo> GetLocalPrinters()
    {
        // ── Попытка 1: Get-CimInstance Win32_Printer (самый надёжный в single-file publish) ──
        try
        {
            var cimResult = GetLocalPrintersCim();
            if (cimResult.Count > 0)
            {
                Serilog.Log.Information("CIM: найдено {Count} принтеров", cimResult.Count);
                return cimResult;
            }
            Serilog.Log.Warning("CIM: вернул 0 принтеров");
        }
        catch (Exception ex)
        {
            Serilog.Log.Warning(ex, "CIM не сработал: {Error}", ex.Message);
        }

        // ── Попытка 2: прямой WMI (может не работать в single-file publish) ──
        try
        {
            var wmiResult = GetLocalPrintersWmi();
            if (wmiResult.Count > 0)
            {
                Serilog.Log.Information("WMI: найдено {Count} принтеров", wmiResult.Count);
                return wmiResult;
            }
            Serilog.Log.Warning("WMI: вернул 0 принтеров");
        }
        catch (Exception ex)
        {
            Serilog.Log.Warning(ex, "WMI не сработал: {Error}", ex.Message);
        }

        // ── Попытка 3: PowerShell Get-Printer (требует модуль) ──
        try
        {
            var psResult = GetLocalPrintersPowerShell();
            if (psResult.Count > 0)
            {
                Serilog.Log.Information("Get-Printer: найдено {Count} принтеров", psResult.Count);
                return psResult;
            }
            Serilog.Log.Warning("Get-Printer: вернул 0 принтеров");
        }
        catch (Exception ex)
        {
            Serilog.Log.Warning(ex, "Get-Printer не сработал: {Error}", ex.Message);
        }

        Serilog.Log.Error("Не удалось получить принтеры ни одним из 3 методов!");
        return [];
    }

    /// <summary>
    /// Get-CimInstance через внешний powershell.exe (Windows PS 5.1 Desktop edition).
    /// CimCmdlets несовместим с Core-режимом встроенного PS SDK — только внешний процесс!
    /// </summary>
    private static IReadOnlyList<PrinterInfo> GetLocalPrintersCim()
    {
        var result = new List<PrinterInfo>();

        // ВАЖНО: используем RunExternalAsync (system powershell.exe),
        // т.к. Get-CimInstance требует CimCmdlets, недоступный в Core-режиме SDK.
        var task = PowerShellEngine.RunExternalAsync(@"
            Get-CimInstance -ClassName Win32_Printer | ForEach-Object {
                $p = $_
                ('{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}' -f `
                    $p.Name, $p.DriverName, $p.PortName, $p.Shared, `
                    $p.ShareName, [int]$p.PrinterStatus, $p.Jobs, `
                    $p.Comment, $p.Location, $p.Default)
            }
        ");
        task.Wait();
        var (success, output, error) = task.Result;

        if (!success || output.Count == 0)
        {
            Serilog.Log.Warning("Get-CimInstance: success={Success}, lines={Lines}, error={Error}",
                success, output.Count, error);
            return result;
        }

        foreach (var line in output)
        {
            var parts = line.Split('|');
            if (parts.Length < 10) continue;

            var name     = parts[0];
            var driver   = parts[1];
            var port     = parts[2];
            var shared   = parts[3].Equals("True", StringComparison.OrdinalIgnoreCase);
            var shareName= parts[4];
            int.TryParse(parts[5], out var statusRaw);
            int.TryParse(parts[6], out var jobCount);
            var comment  = parts[7];
            var location = parts[8];
            var isDefault= parts[9].Equals("True", StringComparison.OrdinalIgnoreCase);

            // Win32_Printer PrinterStatus: 1=Other, 2=Unknown, 3=Idle, 4=Printing, 5=Warmup,
            //   6=StoppedPrinting, 7=Offline
            var status = statusRaw switch
            {
                3 => PrinterStatus.Ready,
                4 => PrinterStatus.Printing,
                7 => PrinterStatus.Offline,
                _ => PrinterStatus.Unknown
            };

            var isNetwork = port.StartsWith(@"\\") ||
                            port.StartsWith("IP_") ||
                            port.StartsWith("WSD");

            result.Add(new PrinterInfo
            {
                Name       = name,
                DriverName = driver,
                PortName   = port,
                IsShared   = shared,
                ShareName  = shareName,
                JobCount   = jobCount,
                Comment    = comment,
                Location   = location,
                ServerName = Environment.MachineName,
                Status     = status,
                IsDefault  = isDefault,
                IsNetwork  = isNetwork,
            });
        }

        return result;
    }

    /// <summary>Прямой WMI запрос — может не работать в single-file publish.</summary>
    private static IReadOnlyList<PrinterInfo> GetLocalPrintersWmi()
    {
        var result = new List<PrinterInfo>();
        using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
        foreach (ManagementObject obj in searcher.Get())
            result.Add(MapPrinter(obj, Environment.MachineName));
        return result;
    }

    /// <summary>PowerShell Get-Printer cmdlet — fallback.</summary>
    private static IReadOnlyList<PrinterInfo> GetLocalPrintersPowerShell()
    {
        var result = new List<PrinterInfo>();

        using var engine = new PowerShellEngine();
        var task = engine.RunAsync(@"
            Get-Printer | ForEach-Object {
                $p = $_
                [PSCustomObject]@{
                    Name         = $p.Name
                    DriverName   = $p.DriverName
                    PortName     = $p.PortName
                    Shared       = $p.Shared
                    ShareName    = $p.ShareName
                    PrinterStatus= [int]$p.PrinterStatus
                    JobCount     = $p.JobCount
                    Comment      = $p.Comment
                    Location     = $p.Location
                    Type         = [int]$p.Type
                }
            } | ForEach-Object {
                ""{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}"" -f $_.Name, $_.DriverName, $_.PortName, $_.Shared, $_.ShareName, $_.PrinterStatus, $_.JobCount, $_.Comment, $_.Location, $_.Type
            }
        ");
        task.Wait();
        var (success, output, error) = task.Result;

        if (!success || output.Count == 0)
            return result;

        foreach (var line in output)
        {
            var parts = line.Split('|');
            if (parts.Length < 10) continue;

            var name     = parts[0];
            var driver   = parts[1];
            var port     = parts[2];
            var shared   = parts[3].Equals("True", StringComparison.OrdinalIgnoreCase);
            var shareName= parts[4];
            int.TryParse(parts[5], out var statusRaw);
            int.TryParse(parts[6], out var jobCount);
            var comment  = parts[7];
            var location = parts[8];
            int.TryParse(parts[9], out var printerType);

            var status = statusRaw switch
            {
                0  => PrinterStatus.Ready,
                1  => PrinterStatus.Paused,
                2  => PrinterStatus.Error,
                3  => PrinterStatus.Deleting,
                4  => PrinterStatus.PaperJam,
                5  => PrinterStatus.PaperOut,
                6  => PrinterStatus.ManualFeed,
                7  => PrinterStatus.PaperProblem,
                8  => PrinterStatus.Offline,
                11 => PrinterStatus.Printing,
                _  => PrinterStatus.Unknown
            };

            var isNetwork = printerType == 1 ||
                            port.StartsWith(@"\\") ||
                            port.StartsWith("IP_") ||
                            port.StartsWith("WSD");

            result.Add(new PrinterInfo
            {
                Name       = name,
                DriverName = driver,
                PortName   = port,
                IsShared   = shared,
                ShareName  = shareName,
                JobCount   = jobCount,
                Comment    = comment,
                Location   = location,
                ServerName = Environment.MachineName,
                Status     = status,
                IsDefault  = false,
                IsNetwork  = isNetwork,
            });
        }

        // Определим принтер по умолчанию
        try
        {
            using var defEngine = new PowerShellEngine();
            var defTask = defEngine.RunAsync(
                "(Get-CimInstance -Class Win32_Printer -Filter \"Default=True\").Name");
            defTask.Wait();
            var (defOk, defOut, _) = defTask.Result;
            if (defOk && defOut.Count > 0)
            {
                var defaultName = defOut[0].Trim();
                var defaultPrinter = result.FirstOrDefault(
                    p => p.Name.Equals(defaultName, StringComparison.OrdinalIgnoreCase));
                if (defaultPrinter != null)
                {
                    var idx = result.IndexOf(defaultPrinter);
                    result[idx] = defaultPrinter with { IsDefault = true };
                }
            }
        }
        catch { /* не критично */ }

        return result;
    }

    /// <summary>Удалённые принтеры через DCOM.</summary>
    private static IReadOnlyList<PrinterInfo> GetRemotePrinters(string machineName,
        string? username, string? password)
    {
        var result = new List<PrinterInfo>();
        try
        {
            var scope = BuildRemoteScope(machineName, username, password);
            using var searcher = new ManagementObjectSearcher(scope,
                new ObjectQuery("SELECT * FROM Win32_Printer"));
            foreach (ManagementObject obj in searcher.Get())
                result.Add(MapPrinter(obj, machineName));

            Serilog.Log.Information("WMI: найдено {Count} принтеров на {Machine}",
                result.Count, machineName);
        }
        catch (Exception ex)
        {
            Serilog.Log.Error(ex, "WMI GetRemotePrinters [{Machine}]", machineName);
            throw; // Пробрасываем — ViewModel покажет ошибку пользователю
        }
        return result;
    }

    private static ManagementScope BuildRemoteScope(string machineName,
        string? username = null, string? password = null)
    {
        var opts = new ConnectionOptions
        {
            Impersonation    = ImpersonationLevel.Impersonate,
            Authentication   = AuthenticationLevel.PacketPrivacy,
            EnablePrivileges = true,
            Timeout          = TimeSpan.FromSeconds(20),
        };
        if (!string.IsNullOrEmpty(username))
        {
            opts.Username = username;
            opts.Password = password;
        }
        var scope = new ManagementScope($@"\\{machineName}\root\cimv2", opts);
        scope.Connect();
        return scope;
    }

    private static PrinterInfo MapPrinter(ManagementObject obj, string machine)
    {
        var statusRaw = safe32(obj["PrinterStatus"]);
        var status = Enum.IsDefined(typeof(PrinterStatus), statusRaw)
            ? (PrinterStatus)statusRaw
            : PrinterStatus.Unknown;

        var portName = str(obj["PortName"]);

        return new PrinterInfo
        {
            Name          = str(obj["Name"]),
            ShareName     = str(obj["ShareName"]),
            PortName      = portName,
            DriverName    = str(obj["DriverName"]),
            Location      = str(obj["Location"]),
            Comment       = str(obj["Comment"]),
            ServerName    = machine,
            Status        = status,
            IsDefault     = (bool)(obj["Default"]  ?? false),
            IsShared      = (bool)(obj["Shared"]   ?? false),
            IsNetwork     = portName.StartsWith(@"\\") ||
                            portName.StartsWith("IP_") ||
                            portName.StartsWith("WSD"),
            JobCount      = safe32(obj["Jobs"]),
            DetectedError = Convert.ToUInt32(obj["DetectedErrorState"] ?? 0u),
        };
    }

    // Helpers
    private static string str(object? v) => v?.ToString() ?? "";
    private static int safe32(object? v) => Convert.ToInt32(v ?? 0);
}

// ── Вспомогательные записи ────────────────────────────────────────────────────

public record PrintJobInfo
{
    public int    JobId      { get; init; }
    public string Document   { get; init; } = "";
    public string Owner      { get; init; } = "";
    public string Status     { get; init; } = "";
    public int    TotalPages { get; init; }
    public long   Size       { get; init; }
}

public record PrinterDriverInfo
{
    public string Name    { get; init; } = "";
    public string Version { get; init; } = "";
    public string InfName { get; init; } = "";
}
