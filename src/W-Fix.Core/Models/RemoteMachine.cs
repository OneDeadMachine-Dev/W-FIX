namespace WFix.Core.Models;

/// <summary>
/// Компьютер из Active Directory, на котором можно выполнить удалённый фикс.
/// </summary>
public record RemoteMachine
{
    public string NetBiosName { get; init; } = "";
    public string Fqdn { get; init; } = "";
    public string OuPath { get; init; } = "";
    public string OperatingSystem { get; init; } = "";
    public DateTime? LastLogon { get; init; }
    public bool IsReachable { get; set; }

    public string DisplayName => string.IsNullOrEmpty(Fqdn) ? NetBiosName : Fqdn;
}
