using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using WFix.Core.Models;

namespace WFix.Core.Services;

/// <summary>
/// Потокобезопасный движок PowerShell.
/// Каждый вызов запускается в собственном Runspace с политикой Bypass (на уровне процесса).
/// Корректно останавливает pipeline/процесс при отмене и ограничивает время выполнения.
/// </summary>
public class PowerShellEngine : IDisposable
{
    private static readonly TimeSpan DefaultExecutionTimeout = TimeSpan.FromMinutes(10);

    private readonly string? _remoteComputer;
    private readonly string? _username;
    private readonly string? _password;
    private bool _disposed;

    public PowerShellEngine(string? remoteComputer = null, string? username = null, string? password = null)
    {
        _remoteComputer = remoteComputer;
        if (!string.IsNullOrWhiteSpace(remoteComputer) && string.IsNullOrWhiteSpace(username) &&
            RemoteCredentialContext.TryResolve(remoteComputer, out var contextualUser, out var contextualPassword))
        {
            _username = contextualUser;
            _password = contextualPassword;
        }
        else
        {
            _username = username;
            _password = password;
        }
    }

    /// <summary>
    /// Выполняет скрипт и возвращает весь вывод строками.
    /// </summary>
    public async Task<PowerShellExecutionResult> RunAsync(
        string script,
        Dictionary<string, object?>? parameters = null,
        CancellationToken ct = default,
        TimeSpan? timeout = null)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentException.ThrowIfNullOrWhiteSpace(script);

        var executionTimeout = ValidateTimeout(timeout);
        using var timeoutCts = new CancellationTokenSource(executionTimeout);
        using var executionCts = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token);

        try
        {
            return await Task.Run(
                () => RunInProcess(script, parameters, executionCts.Token),
                CancellationToken.None);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw new OperationCanceledException(ct);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
        {
            var message = $"PowerShell превысил таймаут {executionTimeout}.";
            return PowerShellExecutionResult.Create(
                [$"[ERROR] {message}"], message, timedOut: true);
        }
    }

    private PowerShellExecutionResult RunInProcess(
        string script,
        Dictionary<string, object?>? parameters,
        CancellationToken ct)
    {
        var lines = new List<string>();
        SecureString? securePassword = null;

        try
        {
            ct.ThrowIfCancellationRequested();
            using var runspace = BuildRunspace();
            runspace.Open();
            using var ps = PowerShell.Create();
            ps.Runspace = runspace;

            if (!string.IsNullOrEmpty(_remoteComputer))
            {
                ps.AddCommand("Invoke-Command")
                    .AddParameter("ComputerName", _remoteComputer)
                    .AddParameter("ScriptBlock", ScriptBlock.Create(script));
                if (!string.IsNullOrEmpty(_username))
                {
                    securePassword = new SecureString();
                    foreach (var character in _password ?? string.Empty)
                        securePassword.AppendChar(character);
                    securePassword.MakeReadOnly();
                    ps.AddParameter("Credential", new PSCredential(_username, securePassword));
                }
            }
            else
            {
                ps.AddScript(script);
            }

            if (parameters != null)
            {
                foreach (var kv in parameters)
                    ps.AddParameter(kv.Key, kv.Value);
            }

            using var cancellationRegistration = ct.Register(() =>
            {
                try { ps.Stop(); }
                catch (ObjectDisposedException) { }
                catch (InvalidOperationException) { }
            });

            var results = ps.Invoke();
            ct.ThrowIfCancellationRequested();

            foreach (var result in results)
                if (result != null) lines.Add(result.ToString() ?? "");

            foreach (var error in ps.Streams.Error)
                lines.Add($"[ERROR] {error}");
            foreach (var warning in ps.Streams.Warning)
                lines.Add($"[WARN] {warning.Message}");
            foreach (var verbose in ps.Streams.Verbose)
                lines.Add($"[VERBOSE] {verbose.Message}");

            return PowerShellExecutionResult.Create(lines);
        }
        catch (PipelineStoppedException) when (ct.IsCancellationRequested)
        {
            throw new OperationCanceledException(ct);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            lines.Add($"[EXCEPTION] {ex.Message}");
            return PowerShellExecutionResult.Create(lines, ex.Message);
        }
        finally
        {
            securePassword?.Dispose();
        }
    }

    private Runspace BuildRunspace()
    {
        var iss = InitialSessionState.CreateDefault();
        iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;

        // Добавляем системный путь к модулям Windows PowerShell,
        // чтобы встроенный SDK мог загружать PrintManagement, DISM и др.
        var systemModulePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            @"WindowsPowerShell\v1.0\Modules");

        if (Directory.Exists(systemModulePath))
        {
            var currentPath = Environment.GetEnvironmentVariable("PSModulePath") ?? "";
            if (!currentPath.Contains(systemModulePath, StringComparison.OrdinalIgnoreCase))
            {
                Environment.SetEnvironmentVariable("PSModulePath",
                    currentPath + ";" + systemModulePath);
            }
        }

        return RunspaceFactory.CreateRunspace(iss);
    }

    /// <summary>
    /// Запуск скрипта через внешний powershell.exe / pwsh.exe.
    /// Используется для cmdlets, требующих модулей Windows (Get-Printer, Get-PrinterDriver, DISM и т.д.),
    /// которые недоступны во встроенном PowerShell SDK.
    /// </summary>
    public static async Task<PowerShellExecutionResult> RunExternalAsync(
        string script,
        CancellationToken ct = default,
        TimeSpan? timeout = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(script);
        ct.ThrowIfCancellationRequested();

        var executionTimeout = ValidateTimeout(timeout);
        var lines = new List<string>();

        try
        {
            var psExe = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "WindowsPowerShell", "v1.0", "powershell.exe");

            if (!File.Exists(psExe))
                psExe = "powershell.exe";

            var utf8Script = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8;\n" + script;
            var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(utf8Script));

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = psExe,
                Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encoded}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };

            using var process = System.Diagnostics.Process.Start(psi);
            if (process == null)
                return PowerShellExecutionResult.Create([], "Не удалось запустить powershell.exe");

            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();

            using var timeoutCts = new CancellationTokenSource(executionTimeout);
            using var executionCts = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token);

            try
            {
                await process.WaitForExitAsync(executionCts.Token);
            }
            catch (OperationCanceledException)
            {
                KillProcessTree(process);
                await process.WaitForExitAsync(CancellationToken.None);

                if (ct.IsCancellationRequested)
                    throw new OperationCanceledException(ct);

                var timeoutMessage = $"PowerShell превысил таймаут {executionTimeout}.";
                return PowerShellExecutionResult.Create(
                    [$"[ERROR] {timeoutMessage}"], timeoutMessage,
                    timedOut: timeoutCts.IsCancellationRequested);
            }

            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            lines.AddRange(stdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries));
            var actionableStderr = IsActionableStandardError(stderr);
            var stderrLines = actionableStderr
                ? stderr.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
                : [];
            lines.AddRange(stderrLines.Select(line => $"[ERROR] {line}"));

            var error = stderrLines.FirstOrDefault();
            if (process.ExitCode != 0 && error is null)
                error = $"Exit code: {process.ExitCode}";

            return PowerShellExecutionResult.Create(lines, error, process.ExitCode);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            lines.Add($"[EXCEPTION] {ex.Message}");
            return PowerShellExecutionResult.Create(lines, ex.Message);
        }
    }

    private static TimeSpan ValidateTimeout(TimeSpan? timeout)
    {
        var value = timeout ?? DefaultExecutionTimeout;
        if (value <= TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(nameof(timeout), "Таймаут должен быть положительным.");
        return value;
    }

    private static void KillProcessTree(System.Diagnostics.Process process)
    {
        try
        {
            if (!process.HasExited)
                process.Kill(entireProcessTree: true);
        }
        catch (InvalidOperationException) { }
        catch (System.ComponentModel.Win32Exception) { }
    }

    private static bool IsActionableStandardError(string stderr)
    {
        if (string.IsNullOrWhiteSpace(stderr))
            return false;

        // Windows PowerShell 5.1 пишет progress-записи запуска модулей в stderr как CLIXML.
        // Это служебный поток, а не ошибка команды. Реальные ErrorRecord помечены S="Error".
        if (stderr.TrimStart().StartsWith("#< CLIXML", StringComparison.OrdinalIgnoreCase))
        {
            return stderr.Contains("S=\"Error\"", StringComparison.OrdinalIgnoreCase) ||
                   stderr.Contains("S='Error'", StringComparison.OrdinalIgnoreCase);
        }

        return true;
    }

    public void Dispose()
    {
        if (!_disposed) _disposed = true;
    }
}
