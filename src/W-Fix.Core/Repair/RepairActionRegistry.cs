using WFix.Core.Abstractions;
using WFix.Core.Services;

namespace WFix.Core.Repair;

public sealed class RepairActionRegistry : IRepairActionRegistry
{
    private readonly IReadOnlyList<IRepairAction> _actions;
    private readonly IReadOnlyDictionary<string, IRepairAction> _byId;

    public RepairActionRegistry(
        FixerRegistry fixerRegistry,
        SystemStateBackupService backupService,
        ICredentialStore credentialStore)
    {
        ArgumentNullException.ThrowIfNull(fixerRegistry);
        ArgumentNullException.ThrowIfNull(backupService);
        ArgumentNullException.ThrowIfNull(credentialStore);

        _actions = fixerRegistry.GetAll()
            .Select(fixer => (IRepairAction)new LegacyFixerRepairAction(fixer, backupService, credentialStore))
            .ToArray();
        _byId = _actions.ToDictionary(action => action.Id, StringComparer.OrdinalIgnoreCase);
    }

    public IReadOnlyList<IRepairAction> GetAll() => _actions;

    public IRepairAction? Get(string actionId) =>
        _byId.GetValueOrDefault(actionId);
}
