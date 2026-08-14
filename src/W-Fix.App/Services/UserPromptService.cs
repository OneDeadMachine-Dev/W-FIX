using System.Windows;

namespace WFix.App.Services;

public interface IUserPromptService
{
    bool Confirm(string message, string title, bool danger = false);
}

public sealed class UserPromptService : IUserPromptService
{
    public bool Confirm(string message, string title, bool danger = false) =>
        MessageBox.Show(
            message,
            title,
            MessageBoxButton.YesNo,
            danger ? MessageBoxImage.Warning : MessageBoxImage.Question) == MessageBoxResult.Yes;
}
