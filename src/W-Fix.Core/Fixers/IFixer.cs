using WFix.Core.Models;

namespace WFix.Core.Fixers;

/// <summary>
/// Базовый интерфейс для всех модулей исправления.
/// </summary>
public interface IFixer
{
    string Name { get; }
    string Description { get; }
    string[] TargetErrorCodes { get; }

    /// <summary>Применить фикс локально или на remoteMachine.</summary>
    Task<FixResult> ApplyAsync(
        PrinterInfo? printer,
        string? remoteMachine,
        IProgress<LogEntry>? progress,
        CancellationToken ct = default);
}

/// <summary>
/// Интерфейс для фиксеров, требующих ввода данных перед запуском.
/// ViewModel должен проверить этот интерфейс и показать диалог.
/// </summary>
public interface IInteractiveFixer
{
    string InputTitle { get; }
    string InputDescription { get; }
    InteractiveInputType InputType { get; }
}

/// <summary>Тип UI, который нужно показать перед запуском фикса.</summary>
public enum InteractiveInputType
{
    /// <summary>Диалог выбора INF / ввода UNC / авто.</summary>
    DriverInstall,
}

/// <summary>
/// Вспомогательный класс для построения списка шагов в фикссерах.
/// </summary>
public abstract class FixerBase : IFixer
{
    public abstract string Name { get; }
    public abstract string Description { get; }
    public abstract string[] TargetErrorCodes { get; }
    public abstract Task<FixResult> ApplyAsync(PrinterInfo? printer, string? remoteMachine, IProgress<LogEntry>? progress, CancellationToken ct);

    protected static LogEntry Info(string msg) => new(Models.LogLevel.Info, msg);
    protected static LogEntry Ok(string msg) => new(Models.LogLevel.Success, msg);
    protected static LogEntry Warn(string msg) => new(Models.LogLevel.Warning, msg);
    protected static LogEntry Err(string msg) => new(Models.LogLevel.Error, msg);
}
