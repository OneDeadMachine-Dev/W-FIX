using WFix.Core.Fixers;

namespace WFix.Core.Services;

/// <summary>
/// Реестр всех доступных фиксеров. 
/// Используйте GetAll() для построения UI-списка, GetByCode() для автоматического подбора.
/// </summary>
public class FixerRegistry
{
    private readonly List<IFixer> _fixers;

    public FixerRegistry()
    {
        _fixers =
        [
            new SpoolerFixer(),
            new Error11bFixer(),
            new Error4005Fixer(),
            new Error709Fixer(),
            new Error02Fixer(),
            new Error7eFixer(),
            new Error7bFixer(),
            new Error8Fixer(),
            new IppFixer(),
            new NetworkFixer(),
            new DriverFixer(),
            new DefaultPrinterFixer(),
        ];
    }

    /// <summary>Все фиксеры для отображения в UI.</summary>
    public IReadOnlyList<IFixer> GetAll() => _fixers.AsReadOnly();

    /// <summary>Фиксеры, применимые к данному коду ошибки.</summary>
    public IReadOnlyList<IFixer> GetByCode(string errorCode)
    {
        var code = errorCode.ToLowerInvariant();
        return _fixers
            .Where(f => f.TargetErrorCodes.Any(c => c.Equals(code, StringComparison.OrdinalIgnoreCase) || code.Contains(c.ToLowerInvariant())))
            .ToList();
    }
}
