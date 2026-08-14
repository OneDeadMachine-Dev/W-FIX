using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Repair;

public sealed class RepairPlanner(IRepairActionRegistry registry) : IRepairPlanner
{
    public RepairPlan CreatePlan(
        TargetDescriptor target,
        IReadOnlyList<DiagnosticFinding> findings,
        RemoteCapabilityReport capabilities)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(findings);
        ArgumentNullException.ThrowIfNull(capabilities);

        var steps = new List<RepairStep>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var finding in findings
                     .OrderByDescending(item => item.Severity)
                     .ThenByDescending(item => item.Confidence))
        {
            foreach (var actionId in finding.RecommendedActionIds)
            {
                var key = $"{actionId}|{finding.PrinterName}";
                if (!seen.Add(key))
                    continue;

                var action = registry.Get(actionId);
                if (action is null || !CapabilitiesSatisfy(capabilities, action.Requirements))
                    continue;

                steps.Add(new RepairStep
                {
                    Id = $"step-{steps.Count + 1:00}",
                    ActionId = action.Id,
                    Title = action.Name,
                    Description = action.Description,
                    PrinterName = finding.PrinterName,
                    Risk = action.Risk,
                    Requirements = action.Requirements,
                    DependsOn = steps.Count == 0 ? [] : [steps[^1].Id]
                });
            }
        }

        return new RepairPlan
        {
            Id = Guid.NewGuid().ToString("N"),
            Target = target,
            Findings = findings,
            Steps = steps
        };
    }

    internal static bool CapabilitiesSatisfy(RemoteCapabilityReport capabilities, RepairRequirement requirements)
    {
        if (requirements.HasFlag(RepairRequirement.Administrator) && !capabilities.IsAdministrator) return false;
        if (requirements.HasFlag(RepairRequirement.WinRm) && !capabilities.WinRmAvailable) return false;
        if (requirements.HasFlag(RepairRequirement.Spooler) && !capabilities.SpoolerAvailable) return false;
        if (requirements.HasFlag(RepairRequirement.TaskScheduler) && !capabilities.TaskSchedulerAvailable) return false;
        if (requirements.HasFlag(RepairRequirement.InteractiveUser) && !capabilities.HasInteractiveUser) return false;
        return true;
    }
}
