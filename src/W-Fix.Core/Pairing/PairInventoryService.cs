using System.Text;
using WFix.Core.Abstractions;
using WFix.Core.Models;
using WFix.Core.Remote;

namespace WFix.Core.Pairing;

public sealed class PairInventoryService(IRemoteSessionFactory sessionFactory) : IPairInventoryService
{
    public async Task<PairEndpointSnapshot> CaptureAsync(
        TargetDescriptor target,
        PairEndpointRole role,
        string peerName,
        string? printerName = null,
        string? shareName = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentException.ThrowIfNullOrWhiteSpace(peerName);
        var script = BuildScript(role, peerName, printerName, shareName);
        await using var session = await sessionFactory.CreateAsync(target, cancellationToken);
        var result = await session.ExecutePowerShellAsync(script, cancellationToken, TimeSpan.FromSeconds(60));
        var dto = PowerShellJson.Deserialize<PairInventoryDto>(result);
        return new PairEndpointSnapshot
        {
            Endpoint = new PairEndpointDescriptor
            {
                Role = role,
                ComputerName = target.ComputerName,
                Fqdn = target.Fqdn,
                IsLocalAgent = target.Source == TargetSource.Local
            },
            OperatingSystem = dto.OperatingSystem ?? "",
            OperatingSystemVersion = dto.OperatingSystemVersion ?? "",
            BuildNumber = dto.BuildNumber,
            UpdateBuildRevision = dto.UpdateBuildRevision,
            DomainJoined = dto.DomainJoined,
            DomainOrWorkgroup = dto.DomainOrWorkgroup ?? "",
            NetworkProfile = dto.NetworkProfile ?? "Unknown",
            Ipv4Addresses = dto.Ipv4Addresses ?? [],
            PeerNameResolved = dto.PeerNameResolved,
            SmbPortReachable = dto.SmbPortReachable,
            RpcEndpointMapperReachable = dto.RpcEndpointMapperReachable,
            SpoolerRunning = dto.SpoolerRunning,
            ServiceStates = dto.ServiceStates ?? new Dictionary<string, string>(),
            NetworkDiscoveryFirewallEnabled = dto.NetworkDiscoveryFirewallEnabled,
            FileAndPrinterSharingFirewallEnabled = dto.FileAndPrinterSharingFirewallEnabled,
            SmbSigningRequired = dto.SmbSigningRequired,
            InsecureGuestLogonsEnabled = dto.InsecureGuestLogonsEnabled,
            HasConflictingSmbConnection = dto.HasConflictingSmbConnection,
            SmbConnectionError = dto.SmbConnectionError,
            RpcOverNamedPipes = dto.RpcOverNamedPipes,
            RpcListenerAllowsNamedPipes = dto.RpcListenerAllowsNamedPipes,
            RpcPrivacyDisabled = dto.RpcPrivacyDisabled,
            RestrictDriverInstallationToAdministrators = dto.RestrictDriverInstallationToAdministrators,
            PrinterName = dto.PrinterName,
            PrinterShareName = dto.PrinterShareName,
            PrinterShared = dto.PrinterShared,
            PrinterAllowsAuthenticatedUsersPrint = dto.PrinterAllowsAuthenticatedUsersPrint,
            PrinterDriverName = dto.PrinterDriverName,
            PrinterDriverVersion = dto.PrinterDriverVersion,
            PrinterConnectionInstalled = dto.PrinterConnectionInstalled,
            RecentErrors = dto.RecentErrors ?? []
        };
    }

    private static string BuildScript(PairEndpointRole role, string peerName, string? printerName, string? shareName)
    {
        static string Encoded(string? value) => Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? ""));
        return $$"""
            $ErrorActionPreference = 'Stop'
            $role = '{{role}}'
            $peerName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encoded(peerName)}}'))
            $requestedPrinter = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encoded(printerName)}}'))
            $requestedShare = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encoded(shareName)}}'))
            $os = Get-CimInstance Win32_OperatingSystem
            $computer = Get-CimInstance Win32_ComputerSystem
            $ubr = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR -ErrorAction SilentlyContinue).UBR
            $profile = Get-NetConnectionProfile -ErrorAction SilentlyContinue | Where-Object { $_.IPv4Connectivity -ne 'Disconnected' } | Select-Object -First 1
            $ipv4 = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {
                $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*'
            } | Select-Object -ExpandProperty IPAddress)
            $resolved = $false
            try { $resolved = @([Net.Dns]::GetHostAddresses($peerName)).Count -gt 0 } catch {}
            $smbReachable = Test-NetConnection -ComputerName $peerName -Port 445 -InformationLevel Quiet -WarningAction SilentlyContinue
            $rpcReachable = Test-NetConnection -ComputerName $peerName -Port 135 -InformationLevel Quiet -WarningAction SilentlyContinue
            $services = [ordered]@{}
            foreach ($name in @('Spooler','LanmanServer','LanmanWorkstation','fdPHost','FDResPub')) {
                $service = Get-Service -Name $name -ErrorAction SilentlyContinue
                $services[$name] = if ($null -eq $service) { 'Missing' } else { [string]$service.Status }
            }
            $enabledRules = @(Get-NetFirewallRule -Enabled True -ErrorAction SilentlyContinue)
            $discoveryFirewall = @($enabledRules | Where-Object {
                $_.DisplayGroup -match 'Network Discovery|Обнаружение сети' -or $_.DisplayName -like 'W-Fix Pair Discovery*'
            }).Count -gt 0
            $printFirewall = @($enabledRules | Where-Object {
                $_.DisplayGroup -match 'File and Printer Sharing|Общий доступ к файлам и принтерам' -or $_.DisplayName -like 'W-Fix Pair Print*'
            }).Count -gt 0
            $smbServer = Get-SmbServerConfiguration -ErrorAction SilentlyContinue
            $smbClient = Get-SmbClientConfiguration -ErrorAction SilentlyContinue
            $conflict = $false
            $smbError = $null
            if ($role -eq 'Client') {
                try { $conflict = @(Get-SmbConnection -ServerName $peerName -ErrorAction SilentlyContinue).Count -gt 1 } catch { $smbError = $_.Exception.Message }
            }
            $rpcPolicy = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC' -ErrorAction SilentlyContinue
            $rpcPrivacy = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' -Name RpcAuthnLevelPrivacyEnabled -ErrorAction SilentlyContinue
            $pointAndPrint = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' -Name RestrictDriverInstallationToAdministrators -ErrorAction SilentlyContinue
            $printer = $null
            if ($role -eq 'Host') {
                $printer = if ($requestedPrinter) { Get-Printer -Name $requestedPrinter -Full -ErrorAction SilentlyContinue } else { Get-Printer -Full -ErrorAction SilentlyContinue | Where-Object Shared | Select-Object -First 1 }
            } elseif ($requestedShare) {
                $connectionName = '\\' + $peerName + '\' + $requestedShare
                $printer = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $connectionName -or $_.ComputerName -eq $peerName -and $_.ShareName -eq $requestedShare } | Select-Object -First 1
            }
            $driverVersion = $null
            $printerAllowsPrint = $false
            if ($null -ne $printer -and $printer.DriverName) {
                $driver = Get-CimInstance Win32_PrinterDriver -ErrorAction SilentlyContinue | Where-Object Name -like ($printer.DriverName + '*') | Select-Object -First 1
                $driverVersion = $driver.Version
            }
            if ($null -ne $printer -and $printer.PermissionSDDL) {
                try {
                    $descriptor = [Security.AccessControl.RawSecurityDescriptor]::new([string]$printer.PermissionSDDL)
                    $allowedSids = @(
                        [Security.Principal.SecurityIdentifier]::new([Security.Principal.WellKnownSidType]::WorldSid, $null).Value,
                        [Security.Principal.SecurityIdentifier]::new([Security.Principal.WellKnownSidType]::AuthenticatedUserSid, $null).Value
                    )
                    $printerAllowsPrint = @($descriptor.DiscretionaryAcl | Where-Object {
                        $_.AceQualifier -eq [Security.AccessControl.AceQualifier]::AccessAllowed -and
                        $allowedSids -contains $_.SecurityIdentifier.Value -and ($_.AccessMask -band 8) -ne 0
                    }).Count -gt 0
                } catch {}
            }
            $events = @()
            foreach ($log in @('Microsoft-Windows-PrintService/Admin','Microsoft-Windows-SMBClient/Connectivity')) {
                try {
                    $events += Get-WinEvent -FilterHashtable @{LogName=$log; Level=1,2,3; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 5 -ErrorAction Stop |
                        ForEach-Object { "$($_.Id):$($_.LevelDisplayName)" }
                } catch {}
            }
            [ordered]@{
                OperatingSystem = $os.Caption
                OperatingSystemVersion = $os.Version
                BuildNumber = [int]$os.BuildNumber
                UpdateBuildRevision = [int]$ubr
                DomainJoined = [bool]$computer.PartOfDomain
                DomainOrWorkgroup = [string]$computer.Domain
                NetworkProfile = if ($null -eq $profile) { 'Unknown' } else { [string]$profile.NetworkCategory }
                Ipv4Addresses = $ipv4
                PeerNameResolved = $resolved
                SmbPortReachable = [bool]$smbReachable
                RpcEndpointMapperReachable = [bool]$rpcReachable
                SpoolerRunning = $services['Spooler'] -eq 'Running'
                ServiceStates = $services
                NetworkDiscoveryFirewallEnabled = $discoveryFirewall
                FileAndPrinterSharingFirewallEnabled = $printFirewall
                SmbSigningRequired = if ($role -eq 'Host') { [bool]$smbServer.RequireSecuritySignature } else { [bool]$smbClient.RequireSecuritySignature }
                InsecureGuestLogonsEnabled = [bool]$smbClient.EnableInsecureGuestLogons
                HasConflictingSmbConnection = $conflict
                SmbConnectionError = $smbError
                RpcOverNamedPipes = [int]$rpcPolicy.RpcUseNamedPipeProtocol -eq 1
                RpcListenerAllowsNamedPipes = (([int]$rpcPolicy.RpcProtocols -band 0x2) -ne 0)
                RpcPrivacyDisabled = ($null -ne $rpcPrivacy) -and ([int]$rpcPrivacy.RpcAuthnLevelPrivacyEnabled -eq 0)
                RestrictDriverInstallationToAdministrators = ($null -eq $pointAndPrint) -or ([int]$pointAndPrint.RestrictDriverInstallationToAdministrators -ne 0)
                PrinterName = $printer.Name
                PrinterShareName = $printer.ShareName
                PrinterShared = [bool]$printer.Shared
                PrinterAllowsAuthenticatedUsersPrint = $printerAllowsPrint
                PrinterDriverName = $printer.DriverName
                PrinterDriverVersion = [string]$driverVersion
                PrinterConnectionInstalled = ($role -eq 'Client') -and ($null -ne $printer)
                RecentErrors = $events
            } | ConvertTo-Json -Depth 6 -Compress
            """;
    }

    private sealed record PairInventoryDto
    {
        public string? OperatingSystem { get; init; }
        public string? OperatingSystemVersion { get; init; }
        public int BuildNumber { get; init; }
        public int UpdateBuildRevision { get; init; }
        public bool DomainJoined { get; init; }
        public string? DomainOrWorkgroup { get; init; }
        public string? NetworkProfile { get; init; }
        public string[]? Ipv4Addresses { get; init; }
        public bool PeerNameResolved { get; init; }
        public bool SmbPortReachable { get; init; }
        public bool RpcEndpointMapperReachable { get; init; }
        public bool SpoolerRunning { get; init; }
        public Dictionary<string, string>? ServiceStates { get; init; }
        public bool NetworkDiscoveryFirewallEnabled { get; init; }
        public bool FileAndPrinterSharingFirewallEnabled { get; init; }
        public bool SmbSigningRequired { get; init; }
        public bool InsecureGuestLogonsEnabled { get; init; }
        public bool HasConflictingSmbConnection { get; init; }
        public string? SmbConnectionError { get; init; }
        public bool RpcOverNamedPipes { get; init; }
        public bool RpcListenerAllowsNamedPipes { get; init; }
        public bool RpcPrivacyDisabled { get; init; }
        public bool RestrictDriverInstallationToAdministrators { get; init; }
        public string? PrinterName { get; init; }
        public string? PrinterShareName { get; init; }
        public bool PrinterShared { get; init; }
        public bool PrinterAllowsAuthenticatedUsersPrint { get; init; }
        public string? PrinterDriverName { get; init; }
        public string? PrinterDriverVersion { get; init; }
        public bool PrinterConnectionInstalled { get; init; }
        public string[]? RecentErrors { get; init; }
    }
}
