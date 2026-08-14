using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Diagnostics;

public interface IDiagnosticService
{
    Task<IReadOnlyList<DiagnosticFinding>> DiagnoseAsync(
        PrinterInventory inventory,
        CancellationToken cancellationToken = default);
}

public sealed class DiagnosticService(IEnumerable<IDiagnosticRule> rules) : IDiagnosticService
{
    private readonly IReadOnlyList<IDiagnosticRule> _rules = rules.ToArray();

    public async Task<IReadOnlyList<DiagnosticFinding>> DiagnoseAsync(
        PrinterInventory inventory,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(inventory);
        var findings = new List<DiagnosticFinding>();
        foreach (var rule in _rules)
        {
            cancellationToken.ThrowIfCancellationRequested();
            findings.AddRange(await rule.EvaluateAsync(inventory, cancellationToken));
        }

        return findings
            .OrderByDescending(finding => finding.Severity)
            .ThenByDescending(finding => finding.Confidence)
            .ToArray();
    }
}
