namespace WFix.Core.Models;

/// <summary>
/// Результат выполнения PowerShell-команды.
/// Дополнительные свойства позволяют отличить ошибку команды от таймаута процесса.
/// </summary>
public sealed record PowerShellExecutionResult(
    bool Success,
    IReadOnlyList<string> Output,
    string? Error)
{
    public int? ExitCode { get; init; }
    public bool TimedOut { get; init; }

    /// <summary>
    /// PowerShell-скрипты W-Fix используют префиксы для машинно-читаемого результата.
    /// Строка [ERROR] должна влиять на итог операции, даже если скрипт перехватил исключение.
    /// </summary>
    public static bool IsErrorLine(string line) =>
        line.StartsWith("[ERROR]", StringComparison.OrdinalIgnoreCase) ||
        line.StartsWith("[EXCEPTION]", StringComparison.OrdinalIgnoreCase);

    public static PowerShellExecutionResult Create(
        IEnumerable<string> output,
        string? error = null,
        int? exitCode = null,
        bool timedOut = false)
    {
        var lines = output.ToList();
        var reportedError = lines.FirstOrDefault(IsErrorLine);
        var effectiveError = error ?? reportedError;
        var success = !timedOut && (exitCode is null || exitCode == 0) && effectiveError is null;

        return new PowerShellExecutionResult(success, lines, effectiveError)
        {
            ExitCode = exitCode,
            TimedOut = timedOut
        };
    }
}
