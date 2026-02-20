using System.Windows;
using System.Windows.Controls;

namespace WFix.App;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void ClearLog_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is ViewModels.MainWindowViewModel vm)
            vm.LiveLog.Clear();
    }

    /// <summary>
    /// Автоскролл лога вниз при добавлении новых записей.
    /// </summary>
    private void LogListBox_ScrollToBottom(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        if (LogListBox.Items.Count > 0)
        {
            LogListBox.ScrollIntoView(LogListBox.Items[^1]);
        }
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        // Подписываемся на изменение коллекции для автоскролла
        if (DataContext is ViewModels.MainWindowViewModel vm)
            vm.LiveLog.CollectionChanged += LogListBox_ScrollToBottom;
    }
}
