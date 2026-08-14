using System.Net;
using System.Net.NetworkInformation;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Remote;

public sealed class RemotePreflightService(IRemoteSessionFactory sessionFactory) : IRemotePreflightService
{
    private const string PreflightScript = """
        $ErrorActionPreference = 'Stop'
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $computer = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $systemDrive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($os.SystemDrive)'" -ErrorAction Stop
        $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
        $scheduler = Get-Service -Name Schedule -ErrorAction SilentlyContinue
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]::new($identity)
        $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $pendingReboot = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
                         (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or
                         ($null -ne (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue))

        [ordered]@{
            CimAvailable = $true
            PowerShellAvailable = $true
            IsAdministrator = $isAdmin
            TaskSchedulerAvailable = ($null -ne $scheduler) -and ($scheduler.Status -eq 'Running')
            SpoolerAvailable = $null -ne $spooler
            InteractiveUser = $computer.UserName
            PendingReboot = $pendingReboot
            OperatingSystem = $os.Caption
            OperatingSystemVersion = $os.Version
            Architecture = $os.OSArchitecture
            FreeSystemDriveBytes = [long]$systemDrive.FreeSpace
        } | ConvertTo-Json -Compress
        """;

    public async Task<RemoteCapabilityReport> CheckAsync(
        TargetDescriptor target,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(target);
        var errors = new List<string>();
        var dnsResolved = target.Source == TargetSource.Local;
        var pingResponded = target.Source == TargetSource.Local;
        long? pingMilliseconds = target.Source == TargetSource.Local ? 0 : null;

        if (target.Source != TargetSource.Local)
        {
            try
            {
                dnsResolved = (await Dns.GetHostAddressesAsync(target.ConnectionName, cancellationToken)).Length > 0;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                errors.Add($"DNS: {ex.Message}");
            }

            try
            {
                using var ping = new Ping();
                var reply = await ping.SendPingAsync(target.ConnectionName, 2000).WaitAsync(cancellationToken);
                pingResponded = reply.Status == IPStatus.Success;
                if (pingResponded)
                    pingMilliseconds = reply.RoundtripTime;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                errors.Add($"Ping: {ex.Message}");
            }
        }

        try
        {
            await using var session = await sessionFactory.CreateAsync(target, cancellationToken);
            var command = await session.ExecutePowerShellAsync(
                PreflightScript,
                cancellationToken,
                TimeSpan.FromSeconds(30));
            var data = PowerShellJson.Deserialize<PreflightDto>(command);
            return new RemoteCapabilityReport
            {
                Target = target,
                DnsResolved = dnsResolved,
                PingResponded = pingResponded,
                PingMilliseconds = pingMilliseconds,
                WinRmAvailable = true,
                CimAvailable = data.CimAvailable,
                PowerShellAvailable = data.PowerShellAvailable,
                IsAdministrator = data.IsAdministrator,
                TaskSchedulerAvailable = data.TaskSchedulerAvailable,
                SpoolerAvailable = data.SpoolerAvailable,
                HasInteractiveUser = !string.IsNullOrWhiteSpace(data.InteractiveUser),
                InteractiveUser = data.InteractiveUser,
                PendingReboot = data.PendingReboot,
                OperatingSystem = data.OperatingSystem,
                OperatingSystemVersion = data.OperatingSystemVersion,
                Architecture = data.Architecture,
                FreeSystemDriveBytes = data.FreeSystemDriveBytes,
                Errors = errors
            };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            errors.Add($"WinRM: {ex.Message}");
            return new RemoteCapabilityReport
            {
                Target = target,
                DnsResolved = dnsResolved,
                PingResponded = pingResponded,
                PingMilliseconds = pingMilliseconds,
                Errors = errors
            };
        }
    }

    private sealed record PreflightDto
    {
        public bool CimAvailable { get; init; }
        public bool PowerShellAvailable { get; init; }
        public bool IsAdministrator { get; init; }
        public bool TaskSchedulerAvailable { get; init; }
        public bool SpoolerAvailable { get; init; }
        public string? InteractiveUser { get; init; }
        public bool PendingReboot { get; init; }
        public string? OperatingSystem { get; init; }
        public string? OperatingSystemVersion { get; init; }
        public string? Architecture { get; init; }
        public long? FreeSystemDriveBytes { get; init; }
    }
}
