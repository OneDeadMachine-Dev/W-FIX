namespace WFix.Core.Models;

/// <summary>
/// Результат применения фиксера.
/// </summary>
public enum FixStatus { Success, Warning, Failed }

public record FixResult
{
    public FixStatus Status { get; init; }
    public string Summary { get; init; } = "";
    public IReadOnlyList<LogEntry> Steps { get; init; } = [];

    public static FixResult Ok(string summary, IEnumerable<LogEntry>? steps = null) =>
        new() { Status = FixStatus.Success, Summary = summary, Steps = steps?.ToList() ?? [] };

    public static FixResult Warn(string summary, IEnumerable<LogEntry>? steps = null) =>
        new() { Status = FixStatus.Warning, Summary = summary, Steps = steps?.ToList() ?? [] };

    public static FixResult Fail(string summary, IEnumerable<LogEntry>? steps = null) =>
        new() { Status = FixStatus.Failed, Summary = summary, Steps = steps?.ToList() ?? [] };
}

public enum LogLevel { Info, Success, Warning, Error }

public record LogEntry(LogLevel Level, string Message, DateTime Timestamp)
{
    public LogEntry(LogLevel level, string message) : this(level, message, DateTime.Now) { }
}
