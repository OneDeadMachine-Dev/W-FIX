using System.Text.Json;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Remote;

public sealed class RemotePrinterInventoryService(IRemoteSessionFactory sessionFactory) : IPrinterInventoryService
{
    private const string InventoryScript = """
        $ErrorActionPreference = 'Stop'
        function Get-PortKind([string]$Name) {
            if ($Name -match '^USB') { return 'USB' }
            if ($Name -match 'IPP|IPPS|https?://') { return 'IPP' }
            if ($Name -match '^WSD') { return 'WSD' }
            if ($Name -match '^\\\\') { return 'UNC' }
            if ($Name -match '^IP_|^TCP') { return 'TCP/IP' }
            return 'Local'
        }

        $os = Get-CimInstance Win32_OperatingSystem
        $ubr = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR -ErrorAction SilentlyContinue).UBR
        $spooler = Get-Service Spooler -ErrorAction SilentlyContinue
        $interactiveUser = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).UserName
        $interactiveDefault = $null
        if (-not [string]::IsNullOrWhiteSpace($interactiveUser)) {
            try {
                $interactiveSid = ([Security.Principal.NTAccount]$interactiveUser).Translate([Security.Principal.SecurityIdentifier]).Value
                $device = (Get-ItemProperty "Registry::HKEY_USERS\$interactiveSid\Software\Microsoft\Windows NT\CurrentVersion\Windows" -Name Device -ErrorAction Stop).Device
                $interactiveDefault = ([string]$device -split ',')[0]
            } catch { }
        }
        $portAddresses = @{}
        try {
            Get-PrinterPort -ErrorAction Stop | ForEach-Object {
                $address = if (-not [string]::IsNullOrWhiteSpace([string]$_.PrinterHostAddress)) { [string]$_.PrinterHostAddress } else { [string]$_.Description }
                $portAddresses[[string]$_.Name] = $address
            }
        } catch { }
        $drivers = @(Get-CimInstance Win32_PrinterDriver -ErrorAction SilentlyContinue | ForEach-Object {
            [ordered]@{
                Name = [string]$_.Name
                Manufacturer = [string]$_.Manufacturer
                Version = [string]$_.DriverVersion
                InfName = [string]$_.InfName
                Architecture = [string]$_.SupportedPlatform
            }
        })
        $printers = @(Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue | ForEach-Object {
            $driverName = [string]$_.DriverName
            $driver = $drivers | Where-Object Name -eq $driverName | Select-Object -First 1
            [ordered]@{
                Name = [string]$_.Name
                DriverName = $driverName
                DriverInfName = [string]$driver.InfName
                DriverManufacturer = [string]$driver.Manufacturer
                DriverVersion = [string]$driver.Version
                PortName = [string]$_.PortName
                PortKind = Get-PortKind ([string]$_.PortName)
                Address = [string]$portAddresses[[string]$_.PortName]
                Status = [int]$_.PrinterStatus
                IsDefault = [bool]$_.Default -or ([string]$_.Name -eq $interactiveDefault)
                IsShared = [bool]$_.Shared
                JobCount = [int]$_.Jobs
                ErrorCodes = @()
            }
        })
        $jobs = @(Get-CimInstance Win32_PrintJob -ErrorAction SilentlyContinue | ForEach-Object {
            [ordered]@{
                Id = [int]$_.JobId
                PrinterName = ([string]$_.Name -split ',')[0]
                Owner = [string]$_.Owner
                Status = [string]$_.Status
                Size = [long]$_.Size
            }
        })
        $hotfixes = @(Get-HotFix -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.HotFixID })
        $pendingReboot = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
                         (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or
                         ($null -ne (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue))

        $pointAndPrint = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' -ErrorAction SilentlyContinue
        $rpc = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC' -ErrorAction SilentlyContinue
        $wppPolicy = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP' -ErrorAction SilentlyContinue
        $wppConfig = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Configuration' -ErrorAction SilentlyContinue
        $wppEnabled = ($wppPolicy.WindowsProtectedPrintMode -eq 1) -or ($wppConfig.WindowsProtectedPrintMode -eq 1)
        $events = @()
        try {
            $events = @(Get-WinEvent -FilterHashtable @{ LogName='Microsoft-Windows-PrintService/Admin'; Level=2 } -MaxEvents 10 -ErrorAction Stop | ForEach-Object {
                '{0:u} Event {1}: {2}' -f $_.TimeCreated, $_.Id, (($_.Message -replace '[\r\n]+',' ') -replace '\s+',' ')
            })
        } catch { }

        [ordered]@{
            OperatingSystem = [string]$os.Caption
            OperatingSystemVersion = [string]$os.Version
            BuildNumber = [int]$os.BuildNumber
            UpdateBuildRevision = [int]$ubr
            Architecture = [string]$os.OSArchitecture
            SpoolerStatus = if ($null -eq $spooler) { 'Missing' } else { [string]$spooler.Status }
            PendingReboot = [bool]$pendingReboot
            WindowsProtectedPrintModeEnabled = [bool]$wppEnabled
            WindowsProtectedPrintModeManaged = ($null -ne $wppPolicy)
            InstalledKnowledgeBaseIds = $hotfixes
            Printers = $printers
            Drivers = $drivers
            Jobs = $jobs
            Policies = [ordered]@{
                RestrictDriverInstallationToAdministrators = [string]$pointAndPrint.RestrictDriverInstallationToAdministrators
                NoWarningNoElevationOnInstall = [string]$pointAndPrint.NoWarningNoElevationOnInstall
                UpdatePromptSettings = [string]$pointAndPrint.UpdatePromptSettings
                RpcUseNamedPipeProtocol = [string]$rpc.RpcUseNamedPipeProtocol
                RpcAuthentication = [string]$rpc.RpcAuthentication
            }
            RecentPrintServiceErrors = $events
        } | ConvertTo-Json -Compress -Depth 7
        """;

    public async Task<PrinterInventory> CaptureAsync(
        TargetDescriptor target,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(target);
        await using var session = await sessionFactory.CreateAsync(target, cancellationToken);
        var result = await session.ExecutePowerShellAsync(
            InventoryScript,
            cancellationToken,
            TimeSpan.FromMinutes(2));
        var data = PowerShellJson.Deserialize<InventoryDto>(result);

        return new PrinterInventory
        {
            Target = target,
            OperatingSystem = data.OperatingSystem ?? string.Empty,
            OperatingSystemVersion = data.OperatingSystemVersion ?? string.Empty,
            BuildNumber = data.BuildNumber,
            UpdateBuildRevision = data.UpdateBuildRevision,
            Architecture = data.Architecture ?? string.Empty,
            SpoolerStatus = data.SpoolerStatus ?? "Unknown",
            PendingReboot = data.PendingReboot,
            WindowsProtectedPrintModeEnabled = data.WindowsProtectedPrintModeEnabled,
            WindowsProtectedPrintModeManaged = data.WindowsProtectedPrintModeManaged,
            InstalledKnowledgeBaseIds = data.InstalledKnowledgeBaseIds ?? [],
            Printers = data.Printers ?? [],
            Drivers = data.Drivers ?? [],
            Jobs = data.Jobs ?? [],
            Policies = ParsePolicies(data.Policies),
            RecentPrintServiceErrors = data.RecentPrintServiceErrors ?? []
        };
    }

    private static IReadOnlyDictionary<string, string?> ParsePolicies(JsonElement policies)
    {
        if (policies.ValueKind != JsonValueKind.Object)
            return new Dictionary<string, string?>();

        return policies.EnumerateObject().ToDictionary(
            property => property.Name,
            property => property.Value.ValueKind == JsonValueKind.Null ? null : property.Value.ToString(),
            StringComparer.OrdinalIgnoreCase);
    }

    private sealed record InventoryDto
    {
        public string? OperatingSystem { get; init; }
        public string? OperatingSystemVersion { get; init; }
        public int BuildNumber { get; init; }
        public int UpdateBuildRevision { get; init; }
        public string? Architecture { get; init; }
        public string? SpoolerStatus { get; init; }
        public bool PendingReboot { get; init; }
        public bool WindowsProtectedPrintModeEnabled { get; init; }
        public bool WindowsProtectedPrintModeManaged { get; init; }
        public string[]? InstalledKnowledgeBaseIds { get; init; }
        public PrinterSnapshot[]? Printers { get; init; }
        public PrinterDriverSnapshot[]? Drivers { get; init; }
        public PrintJobSnapshot[]? Jobs { get; init; }
        public JsonElement Policies { get; init; }
        public string[]? RecentPrintServiceErrors { get; init; }
    }
}
