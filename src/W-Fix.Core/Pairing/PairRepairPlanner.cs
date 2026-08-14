using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class PairRepairPlanner : IPairRepairPlanner
{
    private static readonly IReadOnlyDictionary<string, ActionDefinition> Definitions =
        new Dictionary<string, ActionDefinition>(StringComparer.Ordinal)
        {
            ["pair.discovery.services"] = new("Запустить службы обнаружения", RepairRisk.Reversible, false, FindingEndpoints),
            ["pair.firewall.discovery"] = new("Разрешить точечное обнаружение", RepairRisk.Reversible, false, FindingEndpoints),
            ["pair.firewall.file-print"] = new("Разрешить SMB/RPC печати на хосте", RepairRisk.Reversible, false, _ => [PairEndpointRole.Host]),
            ["pair.spooler.start"] = new("Запустить диспетчер печати хоста", RepairRisk.Reversible, false, _ => [PairEndpointRole.Host]),
            ["pair.smb.clear-conflict"] = new("Закрыть конфликтующий SMB-сеанс", RepairRisk.Irreversible, false, _ => [PairEndpointRole.Client]),
            ["pair.printer.share"] = new("Опубликовать очередь хоста", RepairRisk.Reversible, false, _ => [PairEndpointRole.Host]),
            ["pair.printer.grant-print"] = new("Восстановить право Print общей очереди", RepairRisk.Reversible, false, _ => [PairEndpointRole.Host]),
            ["pair.printer.connect"] = new("Подключить общую очередь", RepairRisk.Reversible, false, _ => [PairEndpointRole.Client]),
            ["pair.rpc.named-pipes"] = new("Включить совместимый RPC over Named Pipes", RepairRisk.Disruptive, true, _ => [PairEndpointRole.Host, PairEndpointRole.Client])
        };

    public PairRepairPlan CreatePlan(
        PairEndpointSnapshot host,
        PairEndpointSnapshot client,
        PairTransportMode transportMode,
        IReadOnlyList<PairDiagnosticFinding> findings,
        bool includeExpertActions = false)
    {
        ArgumentNullException.ThrowIfNull(host);
        ArgumentNullException.ThrowIfNull(client);
        ArgumentNullException.ThrowIfNull(findings);
        var printerName = host.PrinterName ?? client.PrinterName;
        var plannedShareName = host.PrinterShareName ?? client.PrinterShareName ??
            (!string.IsNullOrWhiteSpace(printerName) ? CreateShareName(printerName) : null);
        var steps = new List<PairRepairStep>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var finding in findings.OrderByDescending(item => item.Severity).ThenByDescending(item => item.Confidence))
        {
            foreach (var actionId in finding.RecommendedActionIds)
            {
                if (!Definitions.TryGetValue(actionId, out var definition) || definition.ExpertOnly && !includeExpertActions)
                    continue;
                foreach (var endpoint in definition.ResolveEndpoints(finding))
                {
                    if (!seen.Add($"{actionId}:{endpoint}"))
                        continue;
                    var parameters = BuildParameters(host, printerName, plannedShareName);
                    steps.Add(new PairRepairStep
                    {
                        Id = $"pair-step-{steps.Count + 1:00}",
                        ActionId = actionId,
                        Endpoint = endpoint,
                        Title = definition.Title,
                        Description = finding.Description,
                        Risk = definition.Risk,
                        ExpertOnly = definition.ExpertOnly,
                        Parameters = parameters,
                        DependsOn = steps.Count == 0 ? [] : [steps[^1].Id]
                    });
                }
            }
        }
        return new PairRepairPlan
        {
            Id = Guid.NewGuid().ToString("N"),
            Host = host.Endpoint,
            Client = client.Endpoint,
            TransportMode = transportMode,
            Findings = findings,
            Steps = steps
        };
    }

    private static IReadOnlyList<PairEndpointRole> FindingEndpoints(PairDiagnosticFinding finding) => finding.AffectedEndpoints;

    private static IReadOnlyDictionary<string, string> BuildParameters(
        PairEndpointSnapshot host,
        string? printerName,
        string? shareName)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["hostName"] = host.Endpoint.ComputerName
        };
        if (!string.IsNullOrWhiteSpace(printerName)) values["printerName"] = printerName;
        if (!string.IsNullOrWhiteSpace(shareName)) values["shareName"] = shareName;
        return values;
    }

    internal static string CreateShareName(string printerName)
    {
        var sanitized = new string(printerName.Where(character => char.IsAsciiLetterOrDigit(character) || character is '-' or '_').Take(48).ToArray());
        return string.IsNullOrWhiteSpace(sanitized) ? "WFixPrinter" : sanitized;
    }

    private sealed record ActionDefinition(
        string Title,
        RepairRisk Risk,
        bool ExpertOnly,
        Func<PairDiagnosticFinding, IReadOnlyList<PairEndpointRole>> ResolveEndpoints);
}
