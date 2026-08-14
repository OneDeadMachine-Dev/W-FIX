using System.Globalization;
using System.Text.RegularExpressions;
using WFix.Core.Models;

namespace WFix.Core.Services;

public static partial class TargetParser
{
    public static IReadOnlyList<TargetDescriptor> ParseManual(string? input, CredentialReference? credential = null)
    {
        if (string.IsNullOrWhiteSpace(input))
            return [];

        var result = new Dictionary<string, TargetDescriptor>(StringComparer.OrdinalIgnoreCase);
        foreach (var raw in input.Split([',', ';', '\r', '\n', '\t', ' '], StringSplitOptions.RemoveEmptyEntries))
        {
            var value = raw.Trim().TrimEnd('.');
            if (!ComputerNameRegex().IsMatch(value))
                continue;

            var shortName = value.Split('.')[0].ToUpper(CultureInfo.InvariantCulture);
            var target = new TargetDescriptor
            {
                Id = shortName,
                ComputerName = shortName,
                Fqdn = value.Contains('.') ? value.ToLowerInvariant() : null,
                Source = TargetSource.Manual,
                Credential = credential
            };

            if (!result.TryGetValue(shortName, out var existing) ||
                string.IsNullOrWhiteSpace(existing.Fqdn) && !string.IsNullOrWhiteSpace(target.Fqdn))
                result[shortName] = target;
        }

        return result.Values.OrderBy(target => target.ConnectionName, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public static TargetDescriptor FromActiveDirectory(RemoteMachine machine, CredentialReference? credential = null)
    {
        ArgumentNullException.ThrowIfNull(machine);
        var connectionName = string.IsNullOrWhiteSpace(machine.Fqdn) ? machine.NetBiosName : machine.Fqdn;
        var shortName = string.IsNullOrWhiteSpace(machine.NetBiosName)
            ? connectionName.Split('.')[0]
            : machine.NetBiosName;
        return new TargetDescriptor
        {
            Id = shortName.ToUpperInvariant(),
            ComputerName = shortName.ToUpperInvariant(),
            Fqdn = string.IsNullOrWhiteSpace(machine.Fqdn) ? null : machine.Fqdn.ToLowerInvariant(),
            OuPath = machine.OuPath,
            Source = TargetSource.ActiveDirectory,
            Credential = credential
        };
    }

    [GeneratedRegex(@"^(?=.{1,253}$)(?![-.])(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)(?:\.(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?))*$", RegexOptions.CultureInvariant)]
    private static partial Regex ComputerNameRegex();
}
