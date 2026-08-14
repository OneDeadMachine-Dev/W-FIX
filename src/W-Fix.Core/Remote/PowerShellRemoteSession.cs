using WFix.Core.Abstractions;
using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Remote;

public sealed class PowerShellRemoteSession : IRemoteSession
{
    private readonly PowerShellEngine _engine;

    public PowerShellRemoteSession(TargetDescriptor target, string? username = null, string? password = null)
    {
        Target = target ?? throw new ArgumentNullException(nameof(target));
        _engine = target.Source == TargetSource.Local
            ? new PowerShellEngine()
            : new PowerShellEngine(target.ConnectionName, username, password);
    }

    public TargetDescriptor Target { get; }

    public async Task<RemoteCommandResult> ExecutePowerShellAsync(
        string script,
        CancellationToken cancellationToken = default,
        TimeSpan? timeout = null)
    {
        var result = await _engine.RunAsync(script, ct: cancellationToken, timeout: timeout);
        return new RemoteCommandResult(result.Success, result.Output, result.Error, result.TimedOut);
    }

    public ValueTask DisposeAsync()
    {
        _engine.Dispose();
        return ValueTask.CompletedTask;
    }
}

public sealed class PowerShellRemoteSessionFactory(ICredentialStore credentialStore) : IRemoteSessionFactory
{
    public async Task<IRemoteSession> CreateAsync(
        TargetDescriptor target,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(target);
        if (target.Credential is null)
            return new PowerShellRemoteSession(target);

        var credential = await credentialStore.ReadAsync(target.Credential, cancellationToken)
                         ?? throw new InvalidOperationException($"Учётные данные '{target.Credential.TargetName}' не найдены.");
        return new PowerShellRemoteSession(target, credential.UserName, credential.Password);
    }
}
