using System.Text;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class PairRepairActionRegistry : IPairRepairActionRegistry
{
    private readonly IReadOnlyList<IPairRepairAction> _actions;
    private readonly IReadOnlyDictionary<string, IPairRepairAction> _byId;

    public PairRepairActionRegistry(IRemoteSessionFactory sessionFactory)
    {
        _actions =
        [
            new BuiltInPairRepairAction("pair.discovery.services", "Запустить службы обнаружения", RepairRisk.Reversible, false, true, sessionFactory),
            new BuiltInPairRepairAction("pair.firewall.discovery", "Разрешить сетевое обнаружение", RepairRisk.Reversible, false, true, sessionFactory),
            new BuiltInPairRepairAction("pair.firewall.file-print", "Разрешить SMB/RPC печати", RepairRisk.Reversible, false, true, sessionFactory),
            new BuiltInPairRepairAction("pair.spooler.start", "Запустить диспетчер печати", RepairRisk.Reversible, false, true, sessionFactory),
            new BuiltInPairRepairAction("pair.smb.clear-conflict", "Закрыть конфликтующий SMB-сеанс", RepairRisk.Irreversible, false, false, sessionFactory),
            new BuiltInPairRepairAction("pair.printer.share", "Опубликовать очередь", RepairRisk.Reversible, false, true, sessionFactory),
            new BuiltInPairRepairAction("pair.printer.connect", "Подключить общую очередь", RepairRisk.Reversible, false, true, sessionFactory),
            new BuiltInPairRepairAction("pair.rpc.named-pipes", "Включить RPC over Named Pipes", RepairRisk.Disruptive, true, true, sessionFactory),
            new BuiltInPairRepairAction("pair.rpc.disable-privacy", "Ослабить RPC privacy", RepairRisk.Disruptive, true, true, sessionFactory),
            new BuiltInPairRepairAction("pair.smb.insecure-guest", "Разрешить insecure guest", RepairRisk.Disruptive, true, true, sessionFactory),
            new BuiltInPairRepairAction("pair.smb.disable-signing", "Отключить обязательную SMB-подпись", RepairRisk.Disruptive, true, true, sessionFactory)
        ];
        _byId = _actions.ToDictionary(action => action.Id, StringComparer.Ordinal);
    }

    public IReadOnlyList<IPairRepairAction> GetAll() => _actions;
    public IPairRepairAction? Get(string actionId) => _byId.GetValueOrDefault(actionId);
}

public sealed class RegistryPairActionDispatcher(IPairRepairActionRegistry registry) : IPairActionDispatcher
{
    public Task<PairActionCheckpoint> PrepareAsync(PairActionContext context, CancellationToken cancellationToken = default) =>
        Resolve(context).PrepareAsync(context, cancellationToken);

    public Task<PairActionResult> ExecuteAsync(PairActionContext context, CancellationToken cancellationToken = default) =>
        Resolve(context).ExecuteAsync(context, cancellationToken);

    public Task<bool> VerifyAsync(PairActionContext context, CancellationToken cancellationToken = default) =>
        Resolve(context).VerifyAsync(context, cancellationToken);

    public Task<PairActionResult> RollbackAsync(PairActionContext context, PairActionCheckpoint checkpoint, CancellationToken cancellationToken = default) =>
        Resolve(context).RollbackAsync(context, checkpoint, cancellationToken);

    private IPairRepairAction Resolve(PairActionContext context) =>
        registry.Get(context.Step.ActionId)
        ?? throw new InvalidOperationException($"Pair action '{context.Step.ActionId}' не зарегистрирован.");
}

internal sealed class BuiltInPairRepairAction(
    string id,
    string name,
    RepairRisk risk,
    bool expertOnly,
    bool isIdempotent,
    IRemoteSessionFactory sessionFactory) : IPairRepairAction
{
    public string Id { get; } = id;
    public string Name { get; } = name;
    public RepairRisk Risk { get; } = risk;
    public bool ExpertOnly { get; } = expertOnly;
    public bool IsIdempotent { get; } = isIdempotent;

    public async Task<PairActionCheckpoint> PrepareAsync(PairActionContext context, CancellationToken cancellationToken = default)
    {
        ValidateContext(context);
        var rawState = await RunForJsonAsync(context.Target, BuildPrepareScript(context), context.Step.Timeout, cancellationToken);
        var snapshotDirectory = Path.Combine(context.RunDirectory, "snapshots");
        Directory.CreateDirectory(snapshotDirectory);
        var safeName = context.Step.ActionId.Replace('.', '_') + "-" + context.Step.Endpoint.ToString().ToLowerInvariant() + ".json";
        var snapshotPath = Path.Combine(snapshotDirectory, safeName);
        await File.WriteAllTextAsync(snapshotPath, rawState, new UTF8Encoding(false), cancellationToken);
        return new PairActionCheckpoint
        {
            ActionId = Id,
            Endpoint = context.Step.Endpoint,
            SnapshotPath = snapshotPath,
            State = new Dictionary<string, string?> { ["json"] = rawState }
        };
    }

    public async Task<PairActionResult> ExecuteAsync(PairActionContext context, CancellationToken cancellationToken = default)
    {
        ValidateContext(context);
        var result = await RunAsync(context.Target, BuildExecuteScript(context), context.Step.Timeout, cancellationToken);
        return new PairActionResult
        {
            Success = result.Success,
            Summary = result.Success ? $"{Name}: выполнено." : result.Error ?? $"{Name}: ошибка.",
            Output = RedactOutput(result.Output),
            RequiresReboot = false
        };
    }

    public async Task<bool> VerifyAsync(PairActionContext context, CancellationToken cancellationToken = default)
    {
        ValidateContext(context);
        var result = await RunAsync(context.Target, BuildVerifyScript(context), context.Step.Timeout, cancellationToken);
        return result.Success && result.Output.Any(line => bool.TryParse(line.Trim(), out var value) && value);
    }

    public async Task<PairActionResult> RollbackAsync(
        PairActionContext context,
        PairActionCheckpoint checkpoint,
        CancellationToken cancellationToken = default)
    {
        ValidateContext(context);
        if (!string.Equals(checkpoint.ActionId, Id, StringComparison.Ordinal) || checkpoint.Endpoint != context.Step.Endpoint)
            throw new InvalidOperationException("Checkpoint не соответствует pairing-действию.");
        if (Id == "pair.smb.clear-conflict")
            return new PairActionResult { Success = false, Summary = "Закрытый SMB-сеанс нельзя восстановить без повторной аутентификации." };
        if (!checkpoint.State.TryGetValue("json", out var state) || string.IsNullOrWhiteSpace(state))
            return new PairActionResult { Success = false, Summary = "Checkpoint не содержит состояния для rollback." };
        var result = await RunAsync(context.Target, BuildRollbackScript(context, state), context.Step.Timeout, cancellationToken);
        return new PairActionResult
        {
            Success = result.Success,
            Summary = result.Success ? $"{Name}: rollback выполнен." : result.Error ?? $"{Name}: rollback не выполнен.",
            Output = RedactOutput(result.Output)
        };
    }

    private static void ValidateContext(PairActionContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        if (!context.Step.ActionId.StartsWith("pair.", StringComparison.Ordinal))
            throw new InvalidOperationException("Разрешены только встроенные pairing-действия.");
        if (context.Step.ExpertOnly && context.Step.Risk < RepairRisk.Disruptive)
            throw new InvalidOperationException("Экспертное действие должно иметь повышенную категорию риска.");
    }

    private string BuildPrepareScript(PairActionContext context)
    {
        var rulePrefix = RulePrefix(context);
        return Id switch
        {
            "pair.discovery.services" => """
                $ErrorActionPreference='Stop'
                $result=[ordered]@{}
                foreach($name in @('fdPHost','FDResPub')) { $svc=Get-CimInstance Win32_Service -Filter "Name='$name'"; $result[$name+'Status']=$svc.State; $result[$name+'StartMode']=$svc.StartMode }
                $result | ConvertTo-Json -Compress
                """,
            "pair.firewall.discovery" or "pair.firewall.file-print" => $$"""
                $ErrorActionPreference='Stop'
                $prefix=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(rulePrefix)}}'))
                [ordered]@{Existing=@(Get-NetFirewallRule -ErrorAction SilentlyContinue | Where-Object DisplayName -Like ($prefix+'*') | Select-Object -ExpandProperty DisplayName)} | ConvertTo-Json -Compress
                """,
            "pair.spooler.start" => """
                $ErrorActionPreference='Stop'
                $svc=Get-CimInstance Win32_Service -Filter "Name='Spooler'"
                [ordered]@{Status=$svc.State;StartMode=$svc.StartMode} | ConvertTo-Json -Compress
                """,
            "pair.smb.clear-conflict" => "[ordered]@{RollbackSupported=$false} | ConvertTo-Json -Compress",
            "pair.printer.share" => $$"""
                $ErrorActionPreference='Stop'
                $name=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(Required(context, "printerName"))}}'))
                $printer=Get-Printer -Name $name
                [ordered]@{Shared=[bool]$printer.Shared;ShareName=[string]$printer.ShareName} | ConvertTo-Json -Compress
                """,
            "pair.printer.connect" => $$"""
                $hostName=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(Required(context, "hostName"))}}'))
                $share=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(Required(context, "shareName"))}}'))
                $connection='\\'+$hostName+'\'+$share
                [ordered]@{Existed=($null -ne (Get-Printer -Name $connection -ErrorAction SilentlyContinue));Connection=$connection} | ConvertTo-Json -Compress
                """,
            "pair.rpc.named-pipes" => RegistrySnapshot("HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows NT\\Printers\\RPC", context.Step.Endpoint == PairEndpointRole.Host ? "RpcProtocols" : "RpcUseNamedPipeProtocol"),
            "pair.rpc.disable-privacy" => RegistrySnapshot("HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Print", "RpcAuthnLevelPrivacyEnabled"),
            "pair.smb.insecure-guest" => "[ordered]@{Value=[bool](Get-SmbClientConfiguration).EnableInsecureGuestLogons} | ConvertTo-Json -Compress",
            "pair.smb.disable-signing" => context.Step.Endpoint == PairEndpointRole.Host
                ? "[ordered]@{Value=[bool](Get-SmbServerConfiguration).RequireSecuritySignature} | ConvertTo-Json -Compress"
                : "[ordered]@{Value=[bool](Get-SmbClientConfiguration).RequireSecuritySignature} | ConvertTo-Json -Compress",
            _ => throw new InvalidOperationException($"Неизвестное pairing-действие: {Id}")
        };
    }

    private string BuildExecuteScript(PairActionContext context)
    {
        var rulePrefix = RulePrefix(context);
        return Id switch
        {
            "pair.discovery.services" => """
                $ErrorActionPreference='Stop'
                foreach($name in @('fdPHost','FDResPub')) { Set-Service -Name $name -StartupType Automatic; Start-Service -Name $name }
                'OK'
                """,
            "pair.firewall.discovery" => $$"""
                $ErrorActionPreference='Stop'
                $prefix=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(rulePrefix)}}'))
                if(-not (Get-NetFirewallRule -DisplayName ($prefix+' UDP') -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName ($prefix+' UDP') -Direction Inbound -Action Allow -Profile Domain,Private -RemoteAddress LocalSubnet -Protocol UDP -LocalPort 1900,3702,5355 | Out-Null }
                if(-not (Get-NetFirewallRule -DisplayName ($prefix+' TCP') -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName ($prefix+' TCP') -Direction Inbound -Action Allow -Profile Domain,Private -RemoteAddress LocalSubnet -Protocol TCP -LocalPort 5357,5358 | Out-Null }
                'OK'
                """,
            "pair.firewall.file-print" => $$"""
                $ErrorActionPreference='Stop'
                $prefix=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(rulePrefix)}}'))
                if(-not (Get-NetFirewallRule -DisplayName ($prefix+' SMB') -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName ($prefix+' SMB') -Direction Inbound -Action Allow -Profile Domain,Private -RemoteAddress LocalSubnet -Protocol TCP -LocalPort 445 | Out-Null }
                if(-not (Get-NetFirewallRule -DisplayName ($prefix+' RPC') -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName ($prefix+' RPC') -Direction Inbound -Action Allow -Profile Domain,Private -RemoteAddress LocalSubnet -Protocol TCP -LocalPort 135 | Out-Null }
                if(-not (Get-NetFirewallRule -DisplayName ($prefix+' Spooler') -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName ($prefix+' Spooler') -Direction Inbound -Action Allow -Profile Domain,Private -RemoteAddress LocalSubnet -Program "$env:SystemRoot\System32\spoolsv.exe" -Protocol TCP | Out-Null }
                'OK'
                """,
            "pair.spooler.start" => "Set-Service Spooler -StartupType Automatic; Start-Service Spooler; 'OK'",
            "pair.smb.clear-conflict" => $$"""
                $ErrorActionPreference='Stop'
                $hostName=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(Required(context, "hostName"))}}'))
                Get-SmbMapping -ErrorAction SilentlyContinue | Where-Object RemotePath -Like ('\\'+$hostName+'\*') | Remove-SmbMapping -Force -UpdateProfile -ErrorAction Stop
                & net.exe use ('\\'+$hostName+'\IPC$') /delete /y 2>$null | Out-Null
                'OK'
                """,
            "pair.printer.share" => $$"""
                $ErrorActionPreference='Stop'
                $name=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(Required(context, "printerName"))}}'))
                $share=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(Required(context, "shareName"))}}'))
                Set-Printer -Name $name -Shared $true -ShareName $share
                'OK'
                """,
            "pair.printer.connect" => $$"""
                $ErrorActionPreference='Stop'
                $hostName=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(Required(context, "hostName"))}}'))
                $share=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Encode(Required(context, "shareName"))}}'))
                $connection='\\'+$hostName+'\'+$share
                if(-not (Get-Printer -Name $connection -ErrorAction SilentlyContinue)) { Add-Printer -ConnectionName $connection }
                'OK'
                """,
            "pair.rpc.named-pipes" => SetRegistryScript("HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows NT\\Printers\\RPC", context.Step.Endpoint == PairEndpointRole.Host ? "RpcProtocols" : "RpcUseNamedPipeProtocol", context.Step.Endpoint == PairEndpointRole.Host ? 7 : 1),
            "pair.rpc.disable-privacy" => SetRegistryScript("HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Print", "RpcAuthnLevelPrivacyEnabled", 0),
            "pair.smb.insecure-guest" => "Set-SmbClientConfiguration -EnableInsecureGuestLogons $true -Force -Confirm:$false; 'OK'",
            "pair.smb.disable-signing" => context.Step.Endpoint == PairEndpointRole.Host
                ? "Set-SmbServerConfiguration -RequireSecuritySignature $false -Force -Confirm:$false; 'OK'"
                : "Set-SmbClientConfiguration -RequireSecuritySignature $false -Force -Confirm:$false; 'OK'",
            _ => throw new InvalidOperationException($"Неизвестное pairing-действие: {Id}")
        };
    }

    private string BuildVerifyScript(PairActionContext context)
    {
        var rulePrefix = RulePrefix(context);
        var encodedPrefix = Encode(rulePrefix);
        var encodedHost = context.Step.Parameters.TryGetValue("hostName", out var hostName) ? Encode(hostName) : "";
        var encodedPrinter = context.Step.Parameters.TryGetValue("printerName", out var printerName) ? Encode(printerName) : "";
        var encodedShare = context.Step.Parameters.TryGetValue("shareName", out var shareName) ? Encode(shareName) : "";
        return Id switch
        {
            "pair.discovery.services" => "[bool]((Get-Service fdPHost).Status -eq 'Running' -and (Get-Service FDResPub).Status -eq 'Running')",
            "pair.firewall.discovery" => "$p=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encodedPrefix + "')); [bool](@(Get-NetFirewallRule -Enabled True | Where-Object DisplayName -Like ($p+'*')).Count -ge 2)",
            "pair.firewall.file-print" => "$p=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encodedPrefix + "')); [bool](@(Get-NetFirewallRule -Enabled True | Where-Object DisplayName -Like ($p+'*')).Count -ge 3)",
            "pair.spooler.start" => "[bool]((Get-Service Spooler).Status -eq 'Running')",
            "pair.smb.clear-conflict" => "$h=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encodedHost + "')); [bool](@(Get-SmbMapping -ErrorAction SilentlyContinue | Where-Object RemotePath -Like ('\\\\'+$h+'\\*')).Count -eq 0)",
            "pair.printer.share" => "$n=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encodedPrinter + "'));$s=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encodedShare + "'));$p=Get-Printer -Name $n;[bool]($p.Shared -and $p.ShareName -eq $s)",
            "pair.printer.connect" => "$h=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encodedHost + "'));$s=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + encodedShare + "'));[bool]($null -ne (Get-Printer -Name ('\\\\'+$h+'\\'+$s) -ErrorAction SilentlyContinue))",
            "pair.rpc.named-pipes" => context.Step.Endpoint == PairEndpointRole.Host
                ? @"[bool](((Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC').RpcProtocols -band 2) -ne 0)"
                : @"[bool]((Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC').RpcUseNamedPipeProtocol -eq 1)",
            "pair.rpc.disable-privacy" => @"[bool]((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print').RpcAuthnLevelPrivacyEnabled -eq 0)",
            "pair.smb.insecure-guest" => "[bool](Get-SmbClientConfiguration).EnableInsecureGuestLogons",
            "pair.smb.disable-signing" => context.Step.Endpoint == PairEndpointRole.Host
                ? "[bool](-not (Get-SmbServerConfiguration).RequireSecuritySignature)"
                : "[bool](-not (Get-SmbClientConfiguration).RequireSecuritySignature)",
            _ => "[bool]$false"
        };
    }

    private string BuildRollbackScript(PairActionContext context, string state)
    {
        var load = "$state=([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + Encode(state) + "')) | ConvertFrom-Json);";
        var rulePrefix = RulePrefix(context);
        return Id switch
        {
            "pair.discovery.services" => load + "foreach($n in @('fdPHost','FDResPub')){$mode=$state.($n+'StartMode');$startup=if($mode -eq 'Auto'){'Automatic'}elseif($mode -eq 'Disabled'){'Disabled'}else{'Manual'};Set-Service $n -StartupType $startup;if($state.($n+'Status') -eq 'Running'){Start-Service $n}else{Stop-Service $n -Force -ErrorAction SilentlyContinue}};'OK'",
            "pair.firewall.discovery" or "pair.firewall.file-print" => load + "$p=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + Encode(rulePrefix) + "'));$existing=@($state.Existing);Get-NetFirewallRule -ErrorAction SilentlyContinue|Where-Object{$_.DisplayName -like ($p+'*') -and $_.DisplayName -notin $existing}|Remove-NetFirewallRule;'OK'",
            "pair.spooler.start" => load + "$startup=if($state.StartMode -eq 'Auto'){'Automatic'}elseif($state.StartMode -eq 'Disabled'){'Disabled'}else{'Manual'};Set-Service Spooler -StartupType $startup;if($state.Status -eq 'Running'){Start-Service Spooler}else{Stop-Service Spooler -Force};'OK'",
            "pair.printer.share" => load + "$n=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + Encode(Required(context, "printerName")) + "'));Set-Printer -Name $n -Shared ([bool]$state.Shared) -ShareName ([string]$state.ShareName);'OK'",
            "pair.printer.connect" => load + "if(-not [bool]$state.Existed){Remove-Printer -Name ([string]$state.Connection) -ErrorAction SilentlyContinue};'OK'",
            "pair.rpc.named-pipes" => load + RestoreRegistryScript("HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows NT\\Printers\\RPC", context.Step.Endpoint == PairEndpointRole.Host ? "RpcProtocols" : "RpcUseNamedPipeProtocol"),
            "pair.rpc.disable-privacy" => load + RestoreRegistryScript("HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Print", "RpcAuthnLevelPrivacyEnabled"),
            "pair.smb.insecure-guest" => load + "Set-SmbClientConfiguration -EnableInsecureGuestLogons ([bool]$state.Value) -Force -Confirm:$false;'OK'",
            "pair.smb.disable-signing" => load + (context.Step.Endpoint == PairEndpointRole.Host
                ? "Set-SmbServerConfiguration -RequireSecuritySignature ([bool]$state.Value) -Force -Confirm:$false;'OK'"
                : "Set-SmbClientConfiguration -RequireSecuritySignature ([bool]$state.Value) -Force -Confirm:$false;'OK'"),
            _ => "throw 'Rollback не реализован.'"
        };
    }

    private static string RegistrySnapshot(string path, string name) =>
        "$path='" + path + "';$name='" + name + "';$item=Get-ItemProperty $path -Name $name -ErrorAction SilentlyContinue;" +
        "[ordered]@{Exists=($null -ne $item);Value=if($null -eq $item){$null}else{$item.$name}} | ConvertTo-Json -Compress";

    private static string SetRegistryScript(string path, string name, int value) =>
        $"New-Item -Path '{path}' -Force | Out-Null;New-ItemProperty -Path '{path}' -Name '{name}' -PropertyType DWord -Value {value} -Force | Out-Null;'OK'";

    private static string RestoreRegistryScript(string path, string name) =>
        $"if([bool]$state.Exists){{New-Item -Path '{path}' -Force|Out-Null;New-ItemProperty -Path '{path}' -Name '{name}' -PropertyType DWord -Value ([int]$state.Value) -Force|Out-Null}}else{{Remove-ItemProperty -Path '{path}' -Name '{name}' -ErrorAction SilentlyContinue}};'OK'";

    private static string RulePrefix(PairActionContext context) =>
        $"W-Fix Pair {(context.Step.ActionId.Contains("discovery", StringComparison.Ordinal) ? "Discovery" : "Print")} {Path.GetFileName(context.RunDirectory)} {context.Step.Endpoint}";

    private static string Required(PairActionContext context, string name) =>
        context.Step.Parameters.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new InvalidOperationException($"Pair action '{context.Step.ActionId}' требует параметр '{name}'.");

    private static string Encode(string value) => Convert.ToBase64String(Encoding.UTF8.GetBytes(value));

    private async Task<RemoteCommandResult> RunAsync(TargetDescriptor target, string script, TimeSpan timeout, CancellationToken cancellationToken)
    {
        await using var session = await sessionFactory.CreateAsync(target, cancellationToken);
        return await session.ExecutePowerShellAsync(script, cancellationToken, timeout);
    }

    private async Task<string> RunForJsonAsync(TargetDescriptor target, string script, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var result = await RunAsync(target, script, timeout, cancellationToken);
        if (!result.Success)
            throw new InvalidOperationException(result.Error ?? "Pair action snapshot failed.");
        return result.Output.FirstOrDefault(line => line.TrimStart().StartsWith('{'))
               ?? throw new InvalidDataException("Pair action snapshot не вернул JSON.");
    }

    private static IReadOnlyList<string> RedactOutput(IReadOnlyList<string> output) =>
        output.Where(line => !line.Contains("password", StringComparison.OrdinalIgnoreCase) &&
                             !line.Contains("credential", StringComparison.OrdinalIgnoreCase)).Take(100).ToArray();
}
