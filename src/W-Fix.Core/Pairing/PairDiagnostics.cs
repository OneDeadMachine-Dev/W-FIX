using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class PairDiagnosticService(IEnumerable<IPairDiagnosticRule> rules) : IPairDiagnosticService
{
    private readonly IReadOnlyList<IPairDiagnosticRule> _rules = rules.ToArray();

    public async Task<IReadOnlyList<PairDiagnosticFinding>> DiagnoseAsync(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        CancellationToken cancellationToken = default)
    {
        var findings = new List<PairDiagnosticFinding>();
        foreach (var rule in _rules)
        {
            cancellationToken.ThrowIfCancellationRequested();
            findings.AddRange(await rule.EvaluateAsync(host, client, cancellationToken));
        }
        return findings.OrderByDescending(finding => finding.Severity).ThenByDescending(finding => finding.Confidence).ToArray();
    }
}

public sealed class PairDiscoveryDiagnosticRule : IPairDiagnosticRule
{
    public string Id => "pair.discovery.disabled";

    public Task<IReadOnlyList<PairDiagnosticFinding>> EvaluateAsync(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var affected = new[] { host, client }.Where(snapshot =>
            !snapshot.PeerNameResolved || !snapshot.NetworkDiscoveryFirewallEnabled ||
            snapshot.ServiceStates.TryGetValue("FDResPub", out var resource) && resource != "Running" ||
            snapshot.ServiceStates.TryGetValue("fdPHost", out var provider) && provider != "Running").ToArray();
        if (affected.Length == 0)
            return Task.FromResult<IReadOnlyList<PairDiagnosticFinding>>([]);
        return Task.FromResult<IReadOnlyList<PairDiagnosticFinding>>([new PairDiagnosticFinding
        {
            RuleId = Id,
            Title = "Компьютеры не готовы к сетевому обнаружению",
            Description = "Разрешение имени, Function Discovery или точечные правила Firewall мешают двум выбранным ПК находить друг друга.",
            Severity = FindingSeverity.Warning,
            Confidence = 0.9,
            AffectedEndpoints = affected.Select(snapshot => snapshot.Endpoint.Role).ToArray(),
            Evidence = affected.Select(snapshot => $"{snapshot.Endpoint.Role}: DNS={snapshot.PeerNameResolved}, DiscoveryFirewall={snapshot.NetworkDiscoveryFirewallEnabled}").ToArray(),
            RecommendedActionIds = ["pair.discovery.services", "pair.firewall.discovery"]
        }]);
    }
}

public sealed class PairSmbDiagnosticRule : IPairDiagnosticRule
{
    public string Id => "pair.smb.connectivity";

    public Task<IReadOnlyList<PairDiagnosticFinding>> EvaluateAsync(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var findings = new List<PairDiagnosticFinding>();
        if (!client.SmbPortReachable || !host.FileAndPrinterSharingFirewallEnabled)
        {
            findings.Add(new PairDiagnosticFinding
            {
                RuleId = Id,
                Title = "SMB хоста недоступен клиенту",
                Description = "Без TCP 445 клиент не сможет аутентифицироваться и открыть общую очередь.",
                Severity = FindingSeverity.Error,
                Confidence = 0.98,
                AffectedEndpoints = [PairEndpointRole.Host],
                Evidence = [$"Client TCP/445={client.SmbPortReachable}", $"Host firewall={host.FileAndPrinterSharingFirewallEnabled}"],
                RecommendedActionIds = ["pair.firewall.file-print"]
            });
        }
        if (client.HasConflictingSmbConnection)
        {
            findings.Add(new PairDiagnosticFinding
            {
                RuleId = "pair.smb.credential-conflict",
                Title = "Обнаружено конфликтующее SMB-подключение",
                Description = "Windows не допускает одновременные подключения к одному серверу с разными учётными данными.",
                Severity = FindingSeverity.Error,
                Confidence = 0.95,
                AffectedEndpoints = [PairEndpointRole.Client],
                Evidence = ["У клиента уже есть несколько SMB-сеансов к выбранному хосту."],
                RecommendedActionIds = ["pair.smb.clear-conflict"]
            });
        }
        if (!string.IsNullOrWhiteSpace(client.SmbConnectionError))
        {
            findings.Add(new PairDiagnosticFinding
            {
                RuleId = "pair.smb.authentication",
                Title = "SMB-аутентификация не завершена",
                Description = "Нужно использовать существующую учётную запись хоста и сохранить её только для выбранного имени ПК.",
                Severity = FindingSeverity.Error,
                Confidence = 0.85,
                AffectedEndpoints = [PairEndpointRole.Client],
                Evidence = [client.SmbConnectionError],
                RecommendedActionIds = []
            });
        }
        return Task.FromResult<IReadOnlyList<PairDiagnosticFinding>>(findings);
    }
}

public sealed class PairRpcDiagnosticRule : IPairDiagnosticRule
{
    public string Id => "pair.rpc.connectivity";

    public Task<IReadOnlyList<PairDiagnosticFinding>> EvaluateAsync(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (client.RpcEndpointMapperReachable)
            return Task.FromResult<IReadOnlyList<PairDiagnosticFinding>>([]);
        return Task.FromResult<IReadOnlyList<PairDiagnosticFinding>>([new PairDiagnosticFinding
        {
            RuleId = Id,
            Title = "Print RPC хоста недоступен",
            Description = "Windows 11 использует RPC over TCP по умолчанию; блокировка Endpoint Mapper или Spooler ломает подключение общей очереди.",
            Severity = FindingSeverity.Error,
            Confidence = 0.95,
            AffectedEndpoints = [PairEndpointRole.Host],
            Evidence = ["Client TCP/135=False", $"Host Spooler={host.SpoolerRunning}"],
            RecommendedActionIds = ["pair.firewall.file-print", "pair.spooler.start"] ,
            OfficialSource = new Uri("https://learn.microsoft.com/en-us/troubleshoot/windows-client/printing/windows-11-rpc-connection-updates-for-print")
        }]);
    }
}

public sealed class PairPrinterShareDiagnosticRule : IPairDiagnosticRule
{
    public string Id => "pair.printer.share";

    public Task<IReadOnlyList<PairDiagnosticFinding>> EvaluateAsync(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var findings = new List<PairDiagnosticFinding>();
        if (!host.PrinterShared || string.IsNullOrWhiteSpace(host.PrinterShareName))
        {
            findings.Add(new PairDiagnosticFinding
            {
                RuleId = Id,
                Title = "Принтер хоста не опубликован",
                Description = "Выбранная локальная очередь должна иметь стабильное ShareName и право Print.",
                Severity = FindingSeverity.Error,
                Confidence = 1,
                AffectedEndpoints = [PairEndpointRole.Host],
                Evidence = [$"Printer={host.PrinterName ?? "не выбран"}", $"Shared={host.PrinterShared}"],
                RecommendedActionIds = ["pair.printer.share"]
            });
        }
        if (!client.PrinterConnectionInstalled)
        {
            findings.Add(new PairDiagnosticFinding
            {
                RuleId = "pair.printer.connection-missing",
                Title = "Общая очередь не установлена на клиенте",
                Description = "После восстановления SMB/RPC W-Fix подключит очередь по имени хоста.",
                Severity = FindingSeverity.Warning,
                Confidence = 0.95,
                AffectedEndpoints = [PairEndpointRole.Client],
                Evidence = [$"Host={host.Endpoint.ComputerName}", $"Share={host.PrinterShareName ?? client.PrinterShareName ?? "не задан"}"],
                RecommendedActionIds = ["pair.printer.connect"]
            });
        }
        return Task.FromResult<IReadOnlyList<PairDiagnosticFinding>>(findings);
    }
}

public sealed class PairRpcCompatibilityDiagnosticRule : IPairDiagnosticRule
{
    public string Id => "pair.rpc.compatibility";

    public Task<IReadOnlyList<PairDiagnosticFinding>> EvaluateAsync(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (host.RpcListenerAllowsNamedPipes == client.RpcOverNamedPipes)
            return Task.FromResult<IReadOnlyList<PairDiagnosticFinding>>([]);
        return Task.FromResult<IReadOnlyList<PairDiagnosticFinding>>([new PairDiagnosticFinding
        {
            RuleId = Id,
            Title = "RPC-транспорт клиента и хоста не согласован",
            Description = "Named Pipes не рекомендуется Microsoft и предлагается только как экспертный совместимый режим с rollback.",
            Severity = FindingSeverity.Warning,
            Confidence = 0.8,
            AffectedEndpoints = [PairEndpointRole.Host, PairEndpointRole.Client],
            Evidence = [$"Host listener Named Pipes={host.RpcListenerAllowsNamedPipes}", $"Client uses Named Pipes={client.RpcOverNamedPipes}"],
            RecommendedActionIds = ["pair.rpc.named-pipes"],
            ExpertOnly = true,
            OfficialSource = new Uri("https://learn.microsoft.com/en-us/troubleshoot/windows-client/printing/windows-11-rpc-connection-updates-for-print")
        }]);
    }
}
