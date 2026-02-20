using System.Net.NetworkInformation;
using System.Net.Sockets;
using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Fixers;

/// <summary>
/// Сетевая диагностика принтера: Ping, DNS-резолюция, проверка портов.
/// Не вносит изменений в систему — только диагностирует.
/// </summary>
public class NetworkFixer : FixerBase
{
    private static readonly int[] PrinterPorts = [9100, 515, 443, 80, 631];

    public override string Name => "Состояние — Сетевая диагностика";
    public override string Description =>
        "Проверяет сетевую доступность принтера: Ping, DNS-резолюция имени, " +
        "проверка портов (9100 RAW, 515 LPD, 443 HTTPS, 631 IPP). Не меняет настройки — только диагностирует.";
    public override string[] TargetErrorCodes => ["network", "offline", "dns", "timeout"];

    public override async Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var steps = new List<LogEntry>();
        void Report(LogEntry e) { steps.Add(e); progress?.Report(e); }

        // Определяем цель диагностики
        var target = ExtractTarget(printer);
        if (string.IsNullOrEmpty(target))
        {
            Report(Warn("Не удалось определить IP/имя принтера из PortName. Укажите адрес вручную."));
            return FixResult.Warn("Нет адреса для диагностики", steps);
        }

        Report(Info($"Диагностика сети для: {target}"));

        // 1. Ping
        Report(Info("Проверка Ping..."));
        var pingResult = await PingAsync(target, ct);
        Report(pingResult.Success
            ? Ok($"Ping успешен: {pingResult.RoundtripTime} мс")
            : Err($"Ping недоступен: {pingResult.Error}"));

        // 2. DNS
        Report(Info("Резолюция DNS..."));
        try
        {
            var addresses = await System.Net.Dns.GetHostAddressesAsync(target, ct);
            Report(Ok($"DNS: {string.Join(", ", addresses.Select(a => a.ToString()))}"));
        }
        catch (Exception ex)
        {
            Report(Warn($"DNS не разрешается: {ex.Message}"));
        }

        // 3. Ports
        Report(Info($"Проверка портов: {string.Join(", ", PrinterPorts)}..."));
        foreach (var port in PrinterPorts)
        {
            if (ct.IsCancellationRequested) break;
            var portOpen = await CheckPortAsync(target, port, ct);
            var portName = port switch
            {
                9100 => "RAW",
                515 => "LPD",
                443 => "HTTPS",
                631 => "IPP",
                80 => "HTTP",
                _ => ""
            };
            Report(portOpen
                ? Ok($"  Порт {port} ({portName}): ОТКРЫТ")
                : Warn($"  Порт {port} ({portName}): закрыт / недоступен"));
        }

        // 4. Traceroute (через PS для удобства)
        if (!string.IsNullOrEmpty(remoteMachine))
        {
            Report(Info("Трассировка маршрута (Test-NetConnection)..."));
            var traceScript = $"Test-NetConnection -ComputerName '{target}' -TraceRoute | Select-Object -ExpandProperty TraceRoute";
            using var engine = new PowerShellEngine(remoteMachine);
            var (_, traceOutput, _) = await engine.RunAsync(traceScript, ct: ct);
            foreach (var hop in traceOutput)
                Report(Info($"  Hop: {hop}"));
        }

        var overallOk = pingResult.Success;
        return overallOk
            ? FixResult.Ok($"Принтер {target} сетево доступен", steps)
            : FixResult.Warn($"Принтер {target} не отвечает на Ping — проверьте сеть или IP", steps);
    }

    private static string? ExtractTarget(PrinterInfo? printer)
    {
        if (printer == null) return null;
        var port = printer.PortName;
        if (string.IsNullOrEmpty(port)) return null;

        // "IP_192.168.1.10" — стандартный формат стандартного TCP/IP порта Windows
        if (port.StartsWith("IP_", StringComparison.OrdinalIgnoreCase))
            return port[3..];

        // UNC "\\server\share" — берём server
        if (port.StartsWith(@"\\"))
        {
            var parts = port.TrimStart('\\').Split('\\');
            return parts.FirstOrDefault();
        }

        // Прямое имя/IP
        if (port.Contains('.') || !port.Contains('\\'))
            return port;

        return printer.ServerName;
    }

    private async Task<(bool Success, long RoundtripTime, string? Error)> PingAsync(string host, CancellationToken ct)
    {
        try
        {
            using var ping = new Ping();
            var reply = await ping.SendPingAsync(host, 2000);
            return reply.Status == IPStatus.Success
                ? (true, reply.RoundtripTime, null)
                : (false, 0, reply.Status.ToString());
        }
        catch (Exception ex)
        {
            return (false, 0, ex.Message);
        }
    }

    private async Task<bool> CheckPortAsync(string host, int port, CancellationToken ct)
    {
        try
        {
            using var client = new TcpClient();
            var connectTask = client.ConnectAsync(host, port, ct).AsTask();
            return await Task.WhenAny(connectTask, Task.Delay(1500, ct)) == connectTask && client.Connected;
        }
        catch
        {
            return false;
        }
    }
}
