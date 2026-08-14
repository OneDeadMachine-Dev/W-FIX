using WFix.Core.Abstractions;
using WFix.Core.Fixers;
using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Repair;

/// <summary>
/// Адаптирует проверенные фиксы v2 к конвейеру v3. Это позволяет менять orchestration,
/// не переписывая одновременно рабочие сценарии ремонта.
/// </summary>
public sealed class LegacyFixerRepairAction : IRepairAction
{
    private readonly IFixer _fixer;
    private readonly SystemStateBackupService _backupService;
    private readonly ICredentialStore _credentialStore;

    public LegacyFixerRepairAction(
        IFixer fixer,
        SystemStateBackupService backupService,
        ICredentialStore credentialStore)
    {
        _fixer = fixer ?? throw new ArgumentNullException(nameof(fixer));
        _backupService = backupService ?? throw new ArgumentNullException(nameof(backupService));
        _credentialStore = credentialStore ?? throw new ArgumentNullException(nameof(credentialStore));
    }

    public string Id => GetActionId(_fixer);
    public string Name => _fixer.Name;
    public string Description => _fixer.Description;
    public RepairRisk Risk => _fixer switch
    {
        // Очистка очереди уничтожает документы, а удаление драйверов/файлов не гарантирует автоматическое восстановление.
        SpoolerFixer or DriverFixer or Error7bFixer or Error02Fixer or Error7eFixer => RepairRisk.Irreversible,
        NetworkFixer => RepairRisk.ReadOnly,
        _ => RepairRisk.Reversible
    };
    public RepairRequirement Requirements => GetRequirements(_fixer);
    public bool IsIdempotent => _fixer is SpoolerFixer or NetworkFixer;
    public IReadOnlyList<string> ErrorCodes => _fixer.TargetErrorCodes;

    public async Task<RepairActionCheckpoint> PrepareAsync(
        RepairActionContext context,
        CancellationToken cancellationToken = default)
    {
        using var credentialScope = await CreateCredentialScopeAsync(context.Target, cancellationToken);
        var backup = await _backupService.CreateAsync(
            _fixer,
            context.Printer,
            GetRemoteName(context.Target),
            cancellationToken);

        if (!backup.Success && !backup.Skipped)
            throw new InvalidOperationException($"Не удалось создать снимок: {backup.Error}");

        return new RepairActionCheckpoint(backup.BackupDirectory, backup);
    }

    public async Task<RepairActionResult> ExecuteAsync(
        RepairActionContext context,
        CancellationToken cancellationToken = default)
    {
        using var credentialScope = await CreateCredentialScopeAsync(context.Target, cancellationToken);
        var result = await _fixer.ApplyAsync(
            context.Printer,
            GetRemoteName(context.Target),
            context.Progress,
            cancellationToken);

        return new RepairActionResult(
            result.Status != FixStatus.Failed,
            result.Summary,
            result.Steps.Select(step => $"[{step.Level}] {step.Message}").ToArray());
    }

    public Task<bool> VerifyAsync(
        RepairActionContext context,
        RepairActionResult result,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(result.Success);
    }

    public async Task<RepairActionResult> RollbackAsync(
        RepairActionContext context,
        RepairActionCheckpoint checkpoint,
        CancellationToken cancellationToken = default)
    {
        using var credentialScope = await CreateCredentialScopeAsync(context.Target, cancellationToken);
        if (checkpoint.State is not SystemStateBackupResult backup || backup.Skipped)
            return new RepairActionResult(false, "Для действия нет обратимого снимка.", []);

        var result = await _backupService.RestoreAsync(backup, cancellationToken);
        return new RepairActionResult(result.Success, result.Success ? "Состояние восстановлено." : result.Error ?? "Ошибка отката.", result.Output);
    }

    public static string GetActionId(IFixer fixer) =>
        $"legacy:{fixer.GetType().Name.Replace("Fixer", "", StringComparison.OrdinalIgnoreCase).ToLowerInvariant()}";

    private static RepairRequirement GetRequirements(IFixer fixer)
    {
        var requirements = RepairRequirement.Administrator | RepairRequirement.WinRm | RepairRequirement.Spooler;
        if (fixer is DefaultPrinterFixer or Error709Fixer)
            requirements |= RepairRequirement.TaskScheduler | RepairRequirement.InteractiveUser;
        return requirements;
    }

    private static string? GetRemoteName(TargetDescriptor target) =>
        target.Source == TargetSource.Local ? null : target.ConnectionName;

    private async Task<IDisposable> CreateCredentialScopeAsync(
        TargetDescriptor target,
        CancellationToken cancellationToken)
    {
        if (target.Credential is null || target.Source == TargetSource.Local)
            return EmptyScope.Instance;

        var credential = await _credentialStore.ReadAsync(target.Credential, cancellationToken)
                         ?? throw new InvalidOperationException($"Учётные данные '{target.Credential.TargetName}' не найдены.");
        return RemoteCredentialContext.Push(target.ConnectionName, credential);
    }

    private sealed class EmptyScope : IDisposable
    {
        public static EmptyScope Instance { get; } = new();
        public void Dispose() { }
    }
}
