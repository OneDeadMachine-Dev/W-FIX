using System.Management.Automation.Language;
using WFix.Core.Services;

namespace WFix.Core.Tests;

public sealed class InteractiveUserPowerShellServiceTests
{
    [Fact]
    public void BuildOrchestrationScript_ProducesValidPowerShellAndInteractiveTask()
    {
        const string userScript = "Write-Output '[OK] user context'";

        var script = InteractiveUserPowerShellService.BuildOrchestrationScript(
            userScript,
            "20260814090000-0123456789abcdef",
            TimeSpan.FromSeconds(45));
        Parser.ParseInput(script, out _, out var parseErrors);

        Assert.Empty(parseErrors);
        Assert.Contains("-LogonType Interactive", script);
        Assert.Contains("-RunLevel Limited", script);
        Assert.Contains("ExecutionTimeLimit ([TimeSpan]::FromSeconds(45))", script);
        Assert.Contains("W-Fix-Interactive-20260814090000-0123456789abcdef", script);
        Assert.DoesNotContain(userScript, script);
    }

    [Fact]
    public async Task RunRemoteAsync_RejectsUnsafeShortTimeoutBeforeConnecting()
    {
        var service = new InteractiveUserPowerShellService();

        var exception = await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            service.RunRemoteAsync(
                "unreachable-test-machine",
                "Write-Output test",
                taskTimeout: TimeSpan.FromSeconds(1)));

        Assert.Equal("taskTimeout", exception.ParamName);
    }

    [Fact]
    public async Task RunRemoteAsync_RejectsExcessiveTimeoutBeforeConnecting()
    {
        var service = new InteractiveUserPowerShellService();

        var exception = await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            service.RunRemoteAsync(
                "unreachable-test-machine",
                "Write-Output test",
                taskTimeout: TimeSpan.FromHours(2)));

        Assert.Equal("taskTimeout", exception.ParamName);
    }
}
