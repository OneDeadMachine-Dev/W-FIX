using WFix.Core.Fixers;
using WFix.Core.Models;
using WFix.Core.Services;

namespace WFix.Core.Tests;

public sealed class SystemStateBackupServiceTests
{
    [Fact]
    public void BuildBackupScript_EscapesPowerShellLiterals_AndGeneratesRestoreScript()
    {
        var plan = new SystemStateBackupPlan
        {
            RegistryValues =
            [
                new RegistryValueBackupTarget(@"HKCU:\Software\O'Brien", "Device")
            ]
        };

        var script = SystemStateBackupService.BuildBackupScript("Fixer's test", plan);

        Assert.Contains(@"Path = 'HKCU:\Software\O''Brien'", script);
        Assert.Contains("$safeFixerName = 'Fixer''s test'", script);
        Assert.Contains("restore.ps1", script);
        Assert.Contains("Export-Clixml", script);
        Assert.DoesNotContain("__VALUE_TARGETS__", script);
        Assert.DoesNotContain("__FIXER_NAME__", script);
    }

    [Fact]
    public void Error4005_DeclaresEveryRegistryValueItChanges()
    {
        var fixer = new Error4005Fixer();

        var plan = fixer.CreateBackupPlan(null);

        Assert.Equal(4, plan.RegistryValues.Count);
        Assert.Empty(plan.RegistryKeys);
        Assert.Empty(plan.RegistryAcls);
    }

    [Fact]
    public void Error709_DeclaresValuesAndRegistryAcl()
    {
        var fixer = new Error709Fixer();

        var plan = fixer.CreateBackupPlan(null);

        Assert.Equal(3, plan.RegistryValues.Count);
        Assert.Single(plan.RegistryAcls);
    }

    [Fact]
    public void DefaultPrinter_DeclaresAllPerUserValuesAffectedByDefaultSelection()
    {
        var fixer = new DefaultPrinterFixer();

        var plan = fixer.CreateBackupPlan(null);

        Assert.Equal(3, plan.RegistryValues.Count);
        Assert.Contains(plan.RegistryValues, target => target.Name == "Device");
        Assert.Contains(plan.RegistryValues, target => target.Name == "UserSelectedDefault");
    }

    [Fact]
    public void Error7e_OnlyBacksUpPrinterSpecificBidiKey_WhenPrinterIsSelected()
    {
        var fixer = new Error7eFixer();

        var withoutPrinter = fixer.CreateBackupPlan(null);
        var withPrinter = fixer.CreateBackupPlan(new PrinterInfo { Name = "Office Printer" });

        Assert.Empty(withoutPrinter.RegistryKeys);
        Assert.Single(withPrinter.RegistryKeys);
        Assert.Contains("Office Printer", withPrinter.RegistryKeys[0].NativePath);
    }
}
