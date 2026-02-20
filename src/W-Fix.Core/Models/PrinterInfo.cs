namespace WFix.Core.Models;

/// <summary>
/// Статус принтера, полученный из WMI Win32_Printer.
/// </summary>
public enum PrinterStatus
{
    Unknown = 0,
    Ready = 3,
    Printing = 4,
    Warming = 5,
    Stopped = 6,
    Offline = 7,
    Paused = 8,
    Error = 9,
    Deleting = 10,
    PaperJam = 11,
    PaperOut = 12,
    ManualFeed = 13,
    PaperProblem = 14,
    NotAvailable = 18,
    UserIntervention = 19,
    TonerLow = 20,
    NoToner = 21,
}

/// <summary>
/// Детальная информация о принтере (локальном или сетевом).
/// </summary>
public record PrinterInfo
{
    public string Name { get; init; } = "";
    public string ShareName { get; init; } = "";
    public string PortName { get; init; } = "";
    public string DriverName { get; init; } = "";
    public string Location { get; init; } = "";
    public string Comment { get; init; } = "";
    public string ServerName { get; init; } = "";

    public PrinterStatus Status { get; init; }
    public bool IsDefault { get; init; }
    public bool IsNetwork { get; init; }
    public bool IsShared { get; init; }

    public int JobCount { get; init; }
    public uint DetectedError { get; init; }  // Win32_Printer.DetectedErrorState

    /// <summary>
    /// Список задокументированных кодов ошибок (hex), обнаруженных при диагностике.
    /// </summary>
    public IReadOnlyList<string> ErrorCodes { get; init; } = [];

    /// <summary>
    /// Если принтер получен через AD-запрос — путь в LDAP.
    /// </summary>
    public string? AdPath { get; init; }

    public string StatusDisplay => Status switch
    {
        PrinterStatus.Ready => "Готов",
        PrinterStatus.Printing => "Печатает",
        PrinterStatus.Offline => "Не в сети",
        PrinterStatus.Error => "Ошибка",
        PrinterStatus.PaperJam => "Замятие бумаги",
        PrinterStatus.PaperOut => "Нет бумаги",
        PrinterStatus.TonerLow => "Мало тонера",
        PrinterStatus.NoToner => "Нет тонера",
        PrinterStatus.Paused => "Пауза",
        PrinterStatus.Stopped => "Остановлен",
        PrinterStatus.UserIntervention => "Требует вмешательства",
        _ => Status.ToString()
    };
}
