using System.Diagnostics;
using WFix.Core.Services;

namespace WFix.Core.Tests;

public class PowerShellEngineTests
{
    [Fact]
    public async Task RunAsync_TreatsErrorMarkerAsFailure()
    {
        using var engine = new PowerShellEngine();

        var result = await engine.RunAsync("Write-Output '[ERROR] embedded'");

        Assert.False(result.Success);
        Assert.Contains("[ERROR] embedded", result.Output);
    }

    [Fact]
    public async Task RunExternalAsync_ReturnsOutputForSuccessfulScript()
    {
        var result = await PowerShellEngine.RunExternalAsync("Write-Output '[OK] test'");

        Assert.True(result.Success, result.Error);
        Assert.Contains("[OK] test", result.Output);
        Assert.Equal(0, result.ExitCode);
    }

    [Fact]
    public async Task RunExternalAsync_TreatsErrorMarkerAsFailure()
    {
        var result = await PowerShellEngine.RunExternalAsync("Write-Output '[ERROR] simulated'");

        Assert.False(result.Success);
        Assert.Contains("[ERROR] simulated", result.Output);
    }

    [Fact]
    public async Task RunExternalAsync_StopsProcessAfterTimeout()
    {
        var stopwatch = Stopwatch.StartNew();

        var result = await PowerShellEngine.RunExternalAsync(
            "Start-Sleep -Seconds 30",
            timeout: TimeSpan.FromMilliseconds(500));

        Assert.False(result.Success);
        Assert.True(result.TimedOut);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10));
    }

    [Fact]
    public async Task RunExternalAsync_ThrowsWhenUserCancels()
    {
        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(500));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PowerShellEngine.RunExternalAsync(
                "Start-Sleep -Seconds 30",
                cts.Token,
                TimeSpan.FromMinutes(1)));
    }
}
