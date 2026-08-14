using System.Net;

namespace WFix.Core.Services;

/// <summary>
/// Передаёт credential через legacy-границу IFixer только в пределах async-потока одной цели.
/// Сопоставление имени компьютера не позволяет параллельному заданию использовать чужой секрет.
/// </summary>
internal static class RemoteCredentialContext
{
    private static readonly AsyncLocal<Frame?> CurrentFrame = new();

    public static IDisposable Push(string computerName, NetworkCredential credential)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(computerName);
        ArgumentNullException.ThrowIfNull(credential);
        var previous = CurrentFrame.Value;
        CurrentFrame.Value = new Frame(computerName, credential.UserName, credential.Password, previous);
        return new Scope(previous);
    }

    public static bool TryResolve(string computerName, out string? userName, out string? password)
    {
        for (var frame = CurrentFrame.Value; frame is not null; frame = frame.Previous)
        {
            if (!frame.ComputerName.Equals(computerName, StringComparison.OrdinalIgnoreCase))
                continue;
            userName = frame.UserName;
            password = frame.Password;
            return true;
        }

        userName = null;
        password = null;
        return false;
    }

    private sealed record Frame(string ComputerName, string UserName, string Password, Frame? Previous);

    private sealed class Scope(Frame? previous) : IDisposable
    {
        private Frame? _previous = previous;

        public void Dispose()
        {
            CurrentFrame.Value = Interlocked.Exchange(ref _previous, null);
        }
    }
}
