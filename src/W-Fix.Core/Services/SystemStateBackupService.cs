using System.Text;
using WFix.Core.Fixers;
using WFix.Core.Models;

namespace WFix.Core.Services;

/// <summary>
/// Создаёт точечный снимок изменяемых значений реестра и готовый restore.ps1.
/// Ошибка снимка не блокирует работу фиксера: решение о продолжении принимает UI.
/// </summary>
public sealed class SystemStateBackupService
{
    private const string BackupMarker = "[BACKUP] ";

    public async Task<SystemStateBackupResult> CreateAsync(
        IFixer fixer,
        PrinterInfo? printer,
        string? remoteMachine,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fixer);

        if (fixer is not ISystemStateChangingFixer stateChangingFixer)
            return new SystemStateBackupResult { Success = true, Skipped = true };

        var plan = stateChangingFixer.CreateBackupPlan(printer);
        if (plan.IsEmpty)
            return new SystemStateBackupResult { Success = true, Skipped = true };

        var execution = await ExecuteAsync(
            BuildBackupScript(fixer.Name, plan), remoteMachine, ct);
        var backupDirectory = execution.Output
            .FirstOrDefault(line => line.StartsWith(BackupMarker, StringComparison.OrdinalIgnoreCase))?
            [BackupMarker.Length..].Trim();

        return new SystemStateBackupResult
        {
            Success = execution.Success && !string.IsNullOrWhiteSpace(backupDirectory),
            BackupDirectory = backupDirectory,
            RemoteMachine = remoteMachine,
            Output = execution.Output,
            Error = execution.Error ?? (backupDirectory is null ? "PowerShell не вернул путь к backup." : null)
        };
    }

    public async Task<PowerShellExecutionResult> RestoreAsync(
        SystemStateBackupResult backup,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(backup);
        if (!backup.Success || string.IsNullOrWhiteSpace(backup.BackupDirectory))
            return PowerShellExecutionResult.Create([], "Некорректный или незавершённый backup.");

        var backupPath = EscapePowerShellLiteral(backup.BackupDirectory);
        var script = """
            $ErrorActionPreference = 'Stop'
            try {
                $backupDirectory = '__BACKUP_PATH__'
                $restoreScript = Join-Path $backupDirectory 'restore.ps1'
                if (-not (Test-Path -LiteralPath $restoreScript)) {
                    throw "Файл восстановления не найден: $restoreScript"
                }

                Write-Output "[INFO] Восстановление состояния из: $backupDirectory"
                & $restoreScript
                if ($LASTEXITCODE -ne 0) {
                    throw "restore.ps1 завершился с кодом $LASTEXITCODE"
                }
            } catch {
                Write-Output "[ERROR] Не удалось восстановить backup: $_"
                exit 1
            }
            """.Replace("__BACKUP_PATH__", backupPath, StringComparison.Ordinal);

        return await ExecuteAsync(script, backup.RemoteMachine, ct);
    }

    internal static string BuildBackupScript(string fixerName, SystemStateBackupPlan plan)
    {
        var template = """
            $ErrorActionPreference = 'Stop'
            try {
                $valueTargets = __VALUE_TARGETS__
                $keyTargets = __KEY_TARGETS__
                $aclTargets = __ACL_TARGETS__

                $backupRoot = Join-Path $env:ProgramData 'W-Fix\Backups'
                $safeFixerName = '__FIXER_NAME__' -replace '[^\p{L}\p{Nd}._-]', '_'
                $backupId = '{0}_{1}_{2}' -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $safeFixerName, ([guid]::NewGuid().ToString('N').Substring(0, 8))
                $backupDirectory = Join-Path $backupRoot $backupId
                New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null

                $valueSnapshots = @()
                foreach ($target in $valueTargets) {
                    $exists = $false
                    $value = $null
                    $kind = $null

                    if (Test-Path -LiteralPath $target.Path) {
                        $key = Get-Item -LiteralPath $target.Path -ErrorAction Stop
                        $exists = $key.GetValueNames() -contains $target.Name
                        if ($exists) {
                            $value = $key.GetValue(
                                $target.Name,
                                $null,
                                [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                            $kind = $key.GetValueKind($target.Name).ToString()
                        }
                    }

                    $valueSnapshots += [pscustomobject]@{
                        Path = $target.Path
                        Name = $target.Name
                        Exists = $exists
                        Value = $value
                        Kind = $kind
                    }
                }

                $keySnapshots = @()
                $keyIndex = 0
                foreach ($target in $keyTargets) {
                    $present = Test-Path -LiteralPath $target.ProviderPath
                    $exportFile = $null
                    if ($present) {
                        $exportFile = "registry_key_$keyIndex.reg"
                        & reg.exe export $target.NativePath (Join-Path $backupDirectory $exportFile) /y | Out-Null
                        if ($LASTEXITCODE -ne 0) {
                            throw "reg export завершился с кодом $LASTEXITCODE для $($target.NativePath)"
                        }
                    }

                    $keySnapshots += [pscustomobject]@{
                        ProviderPath = $target.ProviderPath
                        NativePath = $target.NativePath
                        Present = $present
                        ExportFile = $exportFile
                    }
                    $keyIndex++
                }

                $aclSnapshots = @()
                foreach ($target in $aclTargets) {
                    $present = Test-Path -LiteralPath $target.Path
                    $sddl = if ($present) { (Get-Acl -LiteralPath $target.Path).Sddl } else { $null }
                    $aclSnapshots += [pscustomobject]@{
                        Path = $target.Path
                        Present = $present
                        Sddl = $sddl
                    }
                }

                [pscustomobject]@{
                    Values = $valueSnapshots
                    Keys = $keySnapshots
                    Acls = $aclSnapshots
                } | Export-Clixml -LiteralPath (Join-Path $backupDirectory 'state.clixml') -Depth 8

                [ordered]@{
                    Fixer = '__FIXER_NAME__'
                    Computer = $env:COMPUTERNAME
                    CreatedAt = (Get-Date).ToString('o')
                    RegistryValueCount = $valueSnapshots.Count
                    RegistryKeyCount = $keySnapshots.Count
                    RegistryAclCount = $aclSnapshots.Count
                } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $backupDirectory 'manifest.json') -Encoding UTF8

                $restoreScript = @'
            $ErrorActionPreference = 'Stop'
            try {
                $snapshot = Import-Clixml -LiteralPath (Join-Path $PSScriptRoot 'state.clixml')

                foreach ($item in $snapshot.Values) {
                    if ($item.Exists) {
                        if (-not (Test-Path -LiteralPath $item.Path)) {
                            New-Item -Path $item.Path -Force | Out-Null
                        }
                        New-ItemProperty -LiteralPath $item.Path -Name $item.Name -Value $item.Value -PropertyType $item.Kind -Force | Out-Null
                        Write-Output "[OK] Восстановлено значение: $($item.Path)\$($item.Name)"
                    } elseif (Test-Path -LiteralPath $item.Path) {
                        Remove-ItemProperty -LiteralPath $item.Path -Name $item.Name -ErrorAction SilentlyContinue
                        Write-Output "[OK] Удалено ранее отсутствовавшее значение: $($item.Path)\$($item.Name)"
                    }
                }

                foreach ($item in $snapshot.Keys) {
                    if ($item.Present -and $item.ExportFile) {
                        & reg.exe import (Join-Path $PSScriptRoot $item.ExportFile) | Out-Null
                        if ($LASTEXITCODE -ne 0) {
                            throw "reg import завершился с кодом $LASTEXITCODE для $($item.NativePath)"
                        }
                        Write-Output "[OK] Восстановлен раздел: $($item.NativePath)"
                    }
                }

                foreach ($item in $snapshot.Acls) {
                    if ($item.Present -and $item.Sddl -and (Test-Path -LiteralPath $item.Path)) {
                        $acl = Get-Acl -LiteralPath $item.Path
                        $acl.SetSecurityDescriptorSddlForm($item.Sddl)
                        Set-Acl -LiteralPath $item.Path -AclObject $acl
                        Write-Output "[OK] Восстановлены права: $($item.Path)"
                    }
                }

                $spooler = Get-Service -Name spooler -ErrorAction SilentlyContinue
                if ($spooler) {
                    Restart-Service -Name spooler -Force -ErrorAction Stop
                    Write-Output "[OK] Print Spooler перезапущен после восстановления"
                }
                Write-Output "[OK] Восстановление завершено"
            } catch {
                Write-Output "[ERROR] Ошибка restore.ps1: $_"
                exit 1
            }
            '@
                Set-Content -LiteralPath (Join-Path $backupDirectory 'restore.ps1') -Value $restoreScript -Encoding UTF8

                Write-Output "[OK] Снимок состояния создан"
                Write-Output "[BACKUP] $backupDirectory"
            } catch {
                Write-Output "[ERROR] Не удалось создать снимок состояния: $_"
                exit 1
            }
            """;

        return template
            .Replace("__VALUE_TARGETS__", BuildValueTargets(plan.RegistryValues), StringComparison.Ordinal)
            .Replace("__KEY_TARGETS__", BuildKeyTargets(plan.RegistryKeys), StringComparison.Ordinal)
            .Replace("__ACL_TARGETS__", BuildAclTargets(plan.RegistryAcls), StringComparison.Ordinal)
            .Replace("__FIXER_NAME__", EscapePowerShellLiteral(fixerName), StringComparison.Ordinal);
    }

    private static async Task<PowerShellExecutionResult> ExecuteAsync(
        string script,
        string? remoteMachine,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(remoteMachine))
            return await PowerShellEngine.RunExternalAsync(script, ct);

        using var engine = new PowerShellEngine(remoteMachine);
        return await engine.RunAsync(script, ct: ct);
    }

    private static string BuildValueTargets(IReadOnlyList<RegistryValueBackupTarget> targets) =>
        BuildObjectArray(targets.Select(target =>
            $"[pscustomobject]@{{ Path = '{EscapePowerShellLiteral(target.Path)}'; Name = '{EscapePowerShellLiteral(target.Name)}' }}"));

    private static string BuildKeyTargets(IReadOnlyList<RegistryKeyBackupTarget> targets) =>
        BuildObjectArray(targets.Select(target =>
            $"[pscustomobject]@{{ ProviderPath = '{EscapePowerShellLiteral(target.ProviderPath)}'; NativePath = '{EscapePowerShellLiteral(target.NativePath)}' }}"));

    private static string BuildAclTargets(IReadOnlyList<RegistryAclBackupTarget> targets) =>
        BuildObjectArray(targets.Select(target =>
            $"[pscustomobject]@{{ Path = '{EscapePowerShellLiteral(target.Path)}' }}"));

    private static string BuildObjectArray(IEnumerable<string> objects)
    {
        var items = objects.ToList();
        return items.Count == 0
            ? "@()"
            : "@(\n                    " + string.Join(",\n                    ", items) + "\n                )";
    }

    private static string EscapePowerShellLiteral(string value) => value.Replace("'", "''", StringComparison.Ordinal);
}
