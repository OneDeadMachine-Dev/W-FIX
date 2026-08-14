using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Remote;

public sealed class RemoteMaintenanceService(IRemoteSessionFactory sessionFactory) : IRemoteMaintenanceService
{
    public async Task<RemoteCommandResult> RestartAsync(
        TargetDescriptor target,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(target);
        await using var session = await sessionFactory.CreateAsync(target, cancellationToken);
        return await session.ExecutePowerShellAsync(
            "Restart-Computer -Force -ErrorAction Stop; Write-Output '[OK] Команда перезагрузки отправлена.'",
            cancellationToken,
            TimeSpan.FromSeconds(30));
    }
}
