using WFix.Core.Models;

namespace WFix.Core.Tests;

public class PowerShellExecutionResultTests
{
    [Fact]
    public void Create_WithErrorMarker_ReturnsFailure()
    {
        var result = PowerShellExecutionResult.Create(["[ERROR] команда завершилась с ошибкой"]);

        Assert.False(result.Success);
        Assert.Equal("[ERROR] команда завершилась с ошибкой", result.Error);
    }

    [Fact]
    public void Create_WithWarningOnly_ReturnsSuccess()
    {
        var result = PowerShellExecutionResult.Create(["[WARN] требуется перезагрузка"]);

        Assert.True(result.Success);
        Assert.Null(result.Error);
    }

    [Fact]
    public void Create_WithNonZeroExitCode_ReturnsFailure()
    {
        var result = PowerShellExecutionResult.Create([], exitCode: 5);

        Assert.False(result.Success);
        Assert.Equal(5, result.ExitCode);
    }
}
