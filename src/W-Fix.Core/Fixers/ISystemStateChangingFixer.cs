using WFix.Core.Models;

namespace WFix.Core.Fixers;

/// <summary>
/// Фиксер декларирует изменяемые значения/разделы реестра, чтобы исполнитель мог
/// автоматически создать снимок перед применением исправления.
/// </summary>
public interface ISystemStateChangingFixer
{
    SystemStateBackupPlan CreateBackupPlan(PrinterInfo? printer);
}
