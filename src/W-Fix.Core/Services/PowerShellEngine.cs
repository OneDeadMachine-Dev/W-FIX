using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;

namespace WFix.Core.Services;

/// <summary>
/// Потокобезопасный движок PowerShell.
/// Каждый вызов запускается в собственном Runspace с политикой Bypass (на уровне процесса).
/// Поддерживает live-стриминг вывода через IAsyncEnumerable.
/// </summary>
public class PowerShellEngine : IDisposable
{
    private readonly string? _remoteComputer;
    private readonly string? _username;
    private readonly string? _password;
    private bool _disposed;

    public PowerShellEngine(string? remoteComputer = null, string? username = null, string? password = null)
    {
        _remoteComputer = remoteComputer;
        _username = username;
        _password = password;
    }

    /// <summary>
    /// Выполняет скрипт и возвращает весь вывод строками.
    /// </summary>
    public async Task<(bool Success, IReadOnlyList<string> Output, string? Error)> RunAsync(
        string script,
        Dictionary<string, object?>? parameters = null,
        CancellationToken ct = default)
    {
        var lines = new List<string>();
        string? errorMsg = null;
        bool success = true;

        await Task.Run(() =>
        {
            try
            {
                using var runspace = BuildRunspace();
                runspace.Open();
                using var ps = PowerShell.Create();
                ps.Runspace = runspace;

                if (!string.IsNullOrEmpty(_remoteComputer))
                {
                    // Оборачиваем в Invoke-Command для удалённого выполнения
                    var sb = new StringBuilder();
                    sb.Append("Invoke-Command -ComputerName '");
                    sb.Append(_remoteComputer.Replace("'", "''"));
                    sb.Append("' -ScriptBlock { ");
                    sb.Append(script);
                    sb.Append(" }");
                    if (!string.IsNullOrEmpty(_username))
                    {
                        sb.Append(" -Credential (New-Object System.Management.Automation.PSCredential('");
                        sb.Append(_username.Replace("'", "''"));
                        sb.Append("', (ConvertTo-SecureString '");
                        sb.Append((_password ?? "").Replace("'", "''"));
                        sb.Append("' -AsPlainText -Force)))");
                    }
                    ps.AddScript(sb.ToString());
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

                var results = ps.Invoke();
                foreach (var r in results)
                    if (r != null) lines.Add(r.ToString() ?? "");

                foreach (var e in ps.Streams.Error)
                {
                    lines.Add($"[ERROR] {e}");
                    success = false;
                    errorMsg ??= e.ToString();
                }
                foreach (var w in ps.Streams.Warning)
                    lines.Add($"[WARN]  {w.Message}");
                foreach (var v in ps.Streams.Verbose)
                    lines.Add($"[VERBOSE] {v.Message}");
            }
            catch (Exception ex) when (!ct.IsCancellationRequested)
            {
                success = false;
                errorMsg = ex.Message;
                lines.Add($"[EXCEPTION] {ex.Message}");
            }
        }, ct);

        return (success, lines, errorMsg);
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
    public static async Task<(bool Success, IReadOnlyList<string> Output, string? Error)> RunExternalAsync(
        string script, CancellationToken ct = default)
    {
        var lines = new List<string>();
        string? errorMsg = null;
        bool success = true;

        await Task.Run(() =>
        {
            try
            {
                // Ищем powershell.exe (Windows PowerShell 5.1 — есть на всех Windows 10/11)
                var psExe = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
                    "WindowsPowerShell", "v1.0", "powershell.exe");

                if (!File.Exists(psExe))
                {
                    // Fallback на PATH
                    psExe = "powershell.exe";
                }

                // Заставляем PS-скрипт отдавать вывод в UTF-8, чтобы мы могли правильно его прочитать с кириллицей
                var utf8Script = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8;\n" + script;

                // Кодируем скрипт в Base64 для безопасной передачи
                var bytes = System.Text.Encoding.Unicode.GetBytes(utf8Script);
                var encoded = Convert.ToBase64String(bytes);

                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = psExe,
                    Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encoded}",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = System.Text.Encoding.UTF8,
                    StandardErrorEncoding = System.Text.Encoding.UTF8,
                };

                using var process = System.Diagnostics.Process.Start(psi);
                if (process == null)
                {
                    success = false;
                    errorMsg = "Не удалось запустить powershell.exe";
                    return;
                }

                var stdout = process.StandardOutput.ReadToEnd();
                var stderr = process.StandardError.ReadToEnd();
                process.WaitForExit(60_000); // 60 сек таймаут

                foreach (var line in stdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
                    lines.Add(line);

                if (process.ExitCode != 0)
                {
                    success = false;
                    errorMsg = stderr.Length > 0 ? stderr.Trim() : $"Exit code: {process.ExitCode}";
                    if (!string.IsNullOrEmpty(stderr))
                        lines.Add($"[ERROR] {stderr.Trim()}");
                }
            }
            catch (Exception ex) when (!ct.IsCancellationRequested)
            {
                success = false;
                errorMsg = ex.Message;
                lines.Add($"[EXCEPTION] {ex.Message}");
            }
        }, ct);

        return (success, lines, errorMsg);
    }

    public void Dispose()
    {
        if (!_disposed) _disposed = true;
    }
}
