using System.DirectoryServices;
using WFix.Core.Models;

namespace WFix.Core.Services;

/// <summary>
/// Интеграция с Active Directory.
/// Graceful деградация: если домен недоступен — все методы возвращают пустые коллекции.
/// </summary>
public class ActiveDirectoryService
{
    public bool IsDomainAvailable { get; private set; }
    public string DomainName { get; private set; } = "";

    public ActiveDirectoryService()
    {
        try
        {
            DomainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
            IsDomainAvailable = !string.IsNullOrEmpty(DomainName);
        }
        catch
        {
            IsDomainAvailable = false;
        }
    }

    /// <summary>
    /// Поиск компьютеров в AD по маске имени и/или OU.
    /// </summary>
    public IReadOnlyList<RemoteMachine> GetDomainComputers(string? nameFilter = null, string? ouFilter = null)
    {
        if (!IsDomainAvailable) return [];

        var result = new List<RemoteMachine>();
        try
        {
            using var entry = string.IsNullOrEmpty(ouFilter)
                ? new DirectoryEntry("LDAP://RootDSE")
                : new DirectoryEntry($"LDAP://{ouFilter}");

            // Для RootDSE — берём defaultNamingContext
            string ldapBase = ouFilter ?? GetDefaultNamingContext();
            using var baseEntry = new DirectoryEntry($"LDAP://{ldapBase}");
            using var searcher = new DirectorySearcher(baseEntry)
            {
                Filter = BuildComputerFilter(nameFilter),
                PageSize = 500,
                SearchScope = SearchScope.Subtree
            };
            searcher.PropertiesToLoad.AddRange(["name", "dNSHostName", "distinguishedName", "operatingSystem", "lastLogonTimestamp"]);

            using var results = searcher.FindAll();
            foreach (SearchResult sr in results)
            {
                var name = sr.Properties["name"].OfType<string>().FirstOrDefault() ?? "";
                var fqdn = sr.Properties["dNSHostName"].OfType<string>().FirstOrDefault() ?? "";
                var dn = sr.Properties["distinguishedName"].OfType<string>().FirstOrDefault() ?? "";
                var os = sr.Properties["operatingSystem"].OfType<string>().FirstOrDefault() ?? "";

                DateTime? lastLogon = null;
                if (sr.Properties["lastLogonTimestamp"].Count > 0)
                {
                    var raw = (long)sr.Properties["lastLogonTimestamp"][0]!;
                    if (raw > 0) lastLogon = DateTime.FromFileTime(raw);
                }

                result.Add(new RemoteMachine
                {
                    NetBiosName = name,
                    Fqdn = fqdn,
                    OuPath = ExtractOu(dn),
                    OperatingSystem = os,
                    LastLogon = lastLogon
                });
            }
        }
        catch (Exception ex)
        {
            Serilog.Log.Warning(ex, "Ошибка запроса AD GetDomainComputers");
        }
        return result;
    }

    /// <summary>
    /// Поиск принтеров, опубликованных в AD (objectClass=printQueue).
    /// </summary>
    public IReadOnlyList<PrinterInfo> GetPublishedPrinters()
    {
        if (!IsDomainAvailable) return [];

        var result = new List<PrinterInfo>();
        try
        {
            var ldapBase = GetDefaultNamingContext();
            using var baseEntry = new DirectoryEntry($"LDAP://{ldapBase}");
            using var searcher = new DirectorySearcher(baseEntry)
            {
                Filter = "(objectClass=printQueue)",
                PageSize = 200,
                SearchScope = SearchScope.Subtree
            };
            searcher.PropertiesToLoad.AddRange(["printerName", "serverName", "location", "driverName", "portName", "distinguishedName"]);

            using var results = searcher.FindAll();
            foreach (SearchResult sr in results)
            {
                var printerName = sr.Properties["printerName"].OfType<string>().FirstOrDefault() ?? "";
                var serverName = sr.Properties["serverName"].OfType<string>().FirstOrDefault() ?? "";
                var location = sr.Properties["location"].OfType<string>().FirstOrDefault() ?? "";
                var driverName = sr.Properties["driverName"].OfType<string>().FirstOrDefault() ?? "";
                var dn = sr.Properties["distinguishedName"].OfType<string>().FirstOrDefault() ?? "";

                result.Add(new PrinterInfo
                {
                    Name = printerName,
                    ServerName = serverName,
                    Location = location,
                    DriverName = driverName,
                    IsNetwork = true,
                    AdPath = dn
                });
            }
        }
        catch (Exception ex)
        {
            Serilog.Log.Warning(ex, "Ошибка запроса AD GetPublishedPrinters");
        }
        return result;
    }

    private string GetDefaultNamingContext()
    {
        try
        {
            using var rootDse = new DirectoryEntry("LDAP://RootDSE");
            return rootDse.Properties["defaultNamingContext"].Value?.ToString() ?? DomainName;
        }
        catch
        {
            return DomainName;
        }
    }

    private static string BuildComputerFilter(string? nameFilter)
    {
        if (string.IsNullOrWhiteSpace(nameFilter))
            return "(&(objectClass=computer)(objectCategory=computer))";
        var escaped = nameFilter.Replace("(", "\\28").Replace(")", "\\29").Replace("*", "\\2a");
        return $"(&(objectClass=computer)(objectCategory=computer)(name=*{escaped}*))";
    }

    private static string ExtractOu(string dn)
    {
        var parts = dn.Split(',');
        return string.Join(",", parts.Skip(1));
    }
}
