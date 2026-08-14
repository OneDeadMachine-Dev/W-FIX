using System.Text;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class WindowsPairFirewallLeaseService(IRemoteSessionFactory sessionFactory) : IPairFirewallLeaseService
{
    private const string RulePrefix = "W-Fix Pair Session ";

    public async Task<IAsyncDisposable> OpenAsync(
        string sessionId,
        int port,
        string executablePath,
        CancellationToken cancellationToken = default)
    {
        if (!Guid.TryParseExact(sessionId, "N", out _))
            throw new ArgumentException("Некорректный Pair Session ID.", nameof(sessionId));
        if (port is < 1 or > 65535)
            throw new ArgumentOutOfRangeException(nameof(port));
        ArgumentException.ThrowIfNullOrWhiteSpace(executablePath);
        var fullPath = Path.GetFullPath(executablePath);
        var ruleName = RulePrefix + sessionId;
        var encodedName = Encode(ruleName);
        var encodedPath = Encode(fullPath);
        var script = $$"""
            $ErrorActionPreference='Stop'
            $public=@(Get-NetConnectionProfile -ErrorAction SilentlyContinue | Where-Object { $_.IPv4Connectivity -ne 'Disconnected' -and $_.NetworkCategory -eq 'Public' })
            if($public.Count -gt 0){ throw 'Live pairing запрещён для Public network profile.' }
            $name=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{encodedName}}'))
            $program=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{encodedPath}}'))
            if(Get-NetFirewallRule -DisplayName $name -ErrorAction SilentlyContinue){ Remove-NetFirewallRule -DisplayName $name }
            New-NetFirewallRule -DisplayName $name -Description 'Temporary W-Fix pairing listener; removed when the session ends.' -Direction Inbound -Action Allow -Profile Domain,Private -RemoteAddress LocalSubnet -Protocol TCP -LocalPort {{port}} -Program $program | Out-Null
            """;
        await ExecuteAsync(script, cancellationToken);
        return new Lease(this, ruleName);
    }

    public Task CleanupStaleAsync(CancellationToken cancellationToken = default) =>
        ExecuteAsync("Get-NetFirewallRule -ErrorAction SilentlyContinue | Where-Object DisplayName -Like 'W-Fix Pair Session *' | Remove-NetFirewallRule", cancellationToken);

    private async Task RemoveAsync(string ruleName)
    {
        var encoded = Encode(ruleName);
        await ExecuteAsync("$n=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encoded + "'));Get-NetFirewallRule -DisplayName $n -ErrorAction SilentlyContinue|Remove-NetFirewallRule", CancellationToken.None);
    }

    private async Task ExecuteAsync(string script, CancellationToken cancellationToken)
    {
        await using var session = await sessionFactory.CreateAsync(TargetDescriptor.Local(), cancellationToken);
        var result = await session.ExecutePowerShellAsync(script, cancellationToken, TimeSpan.FromSeconds(20));
        if (!result.Success)
            throw new InvalidOperationException(result.Error ?? "Не удалось изменить временное pairing-правило Firewall.");
    }

    private static string Encode(string value) => Convert.ToBase64String(Encoding.UTF8.GetBytes(value));

    private sealed class Lease(WindowsPairFirewallLeaseService owner, string ruleName) : IAsyncDisposable
    {
        private int _disposed;

        public async ValueTask DisposeAsync()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 0)
                await owner.RemoveAsync(ruleName);
        }
    }
}
