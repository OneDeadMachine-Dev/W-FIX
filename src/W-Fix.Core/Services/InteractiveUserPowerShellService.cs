using System.Text;
using WFix.Core.Models;

namespace WFix.Core.Services;

/// <summary>
/// Выполняет PowerShell на удалённой машине в уже существующем интерактивном сеансе пользователя.
/// WinRM работает в административном сервисном контексте и не подходит для per-user операций печати.
/// </summary>
public sealed class InteractiveUserPowerShellService
{
    private static readonly TimeSpan DefaultTaskTimeout = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan MaximumTaskTimeout = TimeSpan.FromHours(1);
    private static readonly TimeSpan CleanupTimeout = TimeSpan.FromSeconds(30);

    public async Task<PowerShellExecutionResult> RunRemoteAsync(
        string remoteMachine,
        string script,
        CancellationToken ct = default,
        TimeSpan? taskTimeout = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(remoteMachine);
        ArgumentException.ThrowIfNullOrWhiteSpace(script);

        var timeout = taskTimeout ?? DefaultTaskTimeout;
        if (timeout < TimeSpan.FromSeconds(5))
            throw new ArgumentOutOfRangeException(nameof(taskTimeout), "Таймаут интерактивной задачи должен быть не меньше 5 секунд.");
        if (timeout > MaximumTaskTimeout)
            throw new ArgumentOutOfRangeException(nameof(taskTimeout), "Таймаут интерактивной задачи не должен превышать 1 час.");

        var executionId = $"{DateTime.UtcNow:yyyyMMddHHmmss}-{Guid.NewGuid():N}";
        var orchestrationScript = BuildOrchestrationScript(script, executionId, timeout);

        try
        {
            using var engine = new PowerShellEngine(remoteMachine);
            return await engine.RunAsync(
                orchestrationScript,
                ct: ct,
                timeout: timeout + TimeSpan.FromMinutes(1));
        }
        catch (OperationCanceledException)
        {
            await TryCleanupRemoteAsync(remoteMachine, executionId);
            throw;
        }
    }

    internal static string BuildOrchestrationScript(
        string userScript,
        string executionId,
        TimeSpan taskTimeout)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(userScript);
        ArgumentException.ThrowIfNullOrWhiteSpace(executionId);

        var payload = Convert.ToBase64String(Encoding.Unicode.GetBytes(userScript));
        var timeoutSeconds = Math.Max(5, (int)Math.Ceiling(taskTimeout.TotalSeconds));

        return $$"""
            $ErrorActionPreference = 'Stop'
            $taskName = 'W-Fix-Interactive-{{executionId}}'
            $scriptPath = Join-Path $env:windir 'Temp\W-Fix-Interactive-{{executionId}}.ps1'
            $outputPath = Join-Path $env:windir 'Temp\W-Fix-Interactive-{{executionId}}.log'
            $taskRegistered = $false
            $failed = $false

            function Set-WFixFileAcl([string]$Path, [System.Security.Principal.SecurityIdentifier]$UserSid, [string]$UserRights) {
                $acl = New-Object System.Security.AccessControl.FileSecurity
                $acl.SetAccessRuleProtection($true, $false)
                $allow = [System.Security.AccessControl.AccessControlType]::Allow
                $systemSid = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
                $administratorsSid = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($systemSid, 'FullControl', $allow)))
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($administratorsSid, 'FullControl', $allow)))
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($UserSid, $UserRights, $allow)))
                Set-Acl -LiteralPath $Path -AclObject $acl
            }

            try {
                $interactiveUser = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).UserName
                if ([string]::IsNullOrWhiteSpace($interactiveUser)) {
                    throw 'На целевом компьютере нет вошедшего интерактивного пользователя.'
                }

                $interactiveSid = ([System.Security.Principal.NTAccount]$interactiveUser).Translate(
                    [System.Security.Principal.SecurityIdentifier])
                Write-Output "[INFO] Интерактивный пользователь: $interactiveUser ($interactiveSid)"

                $userScript = [System.Text.Encoding]::Unicode.GetString(
                    [System.Convert]::FromBase64String('{{payload}}'))
                $runner = @"
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            `$ErrorActionPreference = 'Stop'
            `$exitCode = 0
            try {
                & {
            `$userScript
                } *>&1 | Out-File -LiteralPath '$outputPath' -Encoding utf8
            } catch {
                "[ERROR] Ошибка интерактивного сценария: `$_" | Out-File -LiteralPath '$outputPath' -Encoding utf8 -Append
                `$exitCode = 1
            }
            exit `$exitCode
            "@

                Set-Content -LiteralPath $scriptPath -Value $runner -Encoding Unicode -Force
                Set-Content -LiteralPath $outputPath -Value '' -Encoding UTF8 -Force
                Set-WFixFileAcl -Path $scriptPath -UserSid $interactiveSid -UserRights 'ReadAndExecute'
                Set-WFixFileAcl -Path $outputPath -UserSid $interactiveSid -UserRights 'Modify'

                $powerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
                $action = New-ScheduledTaskAction -Execute $powerShellExe -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""
                $principal = New-ScheduledTaskPrincipal -UserId $interactiveUser -LogonType Interactive -RunLevel Limited
                $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddHours(1)
                $settings = New-ScheduledTaskSettingsSet `
                    -ExecutionTimeLimit ([TimeSpan]::FromSeconds({{timeoutSeconds}})) `
                    -AllowStartIfOnBatteries `
                    -DontStopIfGoingOnBatteries

                Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal `
                    -Trigger $trigger -Settings $settings -Force | Out-Null
                $taskRegistered = $true

                $startedAt = Get-Date
                $deadline = $startedAt.AddSeconds({{timeoutSeconds + 10}})
                $hasStarted = $false
                Start-ScheduledTask -TaskName $taskName

                do {
                    Start-Sleep -Milliseconds 250
                    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
                    $taskInfo = Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction Stop
                    if ($task.State -in @('Running', 'Queued') -or $taskInfo.LastRunTime -ge $startedAt.AddSeconds(-2)) {
                        $hasStarted = $true
                    }
                    $completed = $hasStarted -and $task.State -notin @('Running', 'Queued')
                } while (-not $completed -and (Get-Date) -lt $deadline)

                if (-not $completed) {
                    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
                    throw 'Интерактивная задача не завершилась за отведённое время.'
                }

                if (Test-Path -LiteralPath $outputPath) {
                    Get-Content -LiteralPath $outputPath -Encoding UTF8 |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                        Write-Output
                }

                $taskInfo = Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction Stop
                if ($taskInfo.LastTaskResult -ne 0) {
                    throw "Интерактивная задача завершилась с кодом $($taskInfo.LastTaskResult)."
                }
                Write-Output '[OK] Интерактивный пользовательский сценарий завершён'
            } catch {
                Write-Output "[ERROR] Не удалось выполнить сценарий в интерактивном сеансе: $_"
                $failed = $true
            } finally {
                if ($taskRegistered) {
                    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
                    try {
                        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
                    } catch {
                        Write-Output "[WARN] Не удалось удалить временную задачу $taskName`: $_"
                    }
                }
                foreach ($temporaryPath in @($scriptPath, $outputPath)) {
                    if (Test-Path -LiteralPath $temporaryPath) {
                        try {
                            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction Stop
                        } catch {
                            Write-Output "[WARN] Не удалось удалить временный файл $temporaryPath`: $_"
                        }
                    }
                }
            }

            if ($failed) { exit 1 }
            """;
    }

    private static async Task TryCleanupRemoteAsync(string remoteMachine, string executionId)
    {
        var cleanupScript = $$"""
            $taskName = 'W-Fix-Interactive-{{executionId}}'
            Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath (Join-Path $env:windir 'Temp\W-Fix-Interactive-{{executionId}}.ps1') -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath (Join-Path $env:windir 'Temp\W-Fix-Interactive-{{executionId}}.log') -Force -ErrorAction SilentlyContinue
            """;

        try
        {
            using var cleanupEngine = new PowerShellEngine(remoteMachine);
            var result = await cleanupEngine.RunAsync(cleanupScript, timeout: CleanupTimeout);
            if (!result.Success)
            {
                Serilog.Log.Warning(
                    "Не удалось очистить интерактивную задачу '{ExecutionId}' на {RemoteMachine}: {Error}",
                    executionId,
                    remoteMachine,
                    result.Error);
            }
        }
        catch (Exception ex)
        {
            // Task Scheduler всё равно остановит процесс по ExecutionTimeLimit.
            Serilog.Log.Warning(
                ex,
                "Исключение при очистке интерактивной задачи '{ExecutionId}' на {RemoteMachine}",
                executionId,
                remoteMachine);
        }
    }
}
