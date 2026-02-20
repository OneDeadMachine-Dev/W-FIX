using System.Windows;
using Microsoft.Win32;
using WFix.Core.Fixers;

namespace WFix.App.Views;

public partial class DriverInstallDialog : Window
{
    /// <summary>Выбранный режим.</summary>
    public DriverFixMode SelectedMode { get; private set; } = DriverFixMode.Inf;

    /// <summary>Путь к INF или UNC (зависит от режима).</summary>
    public string SelectedPath { get; private set; } = "";

    public DriverInstallDialog()
    {
        InitializeComponent();
    }

    private void ModeChanged(object sender, RoutedEventArgs e)
    {
        if (InfPanel == null) return; // designer-time guard

        InfPanel.Visibility = RbInf.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        UncPanel.Visibility = RbUnc.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        AutoPanel.Visibility = RbAuto.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
    }

    private void BrowseInf_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Выберите INF-файл драйвера принтера",
            Filter = "INF-файлы (*.inf)|*.inf|Все файлы (*.*)|*.*",
            CheckFileExists = true,
        };
        if (dlg.ShowDialog(this) == true)
        {
            TxtInfPath.Text = dlg.FileName;
        }
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        if (RbInf.IsChecked == true)
        {
            SelectedMode = DriverFixMode.Inf;
            SelectedPath = TxtInfPath.Text.Trim();
            if (string.IsNullOrEmpty(SelectedPath))
            {
                MessageBox.Show(this, "Укажите путь к INF-файлу.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        else if (RbUnc.IsChecked == true)
        {
            SelectedMode = DriverFixMode.Unc;
            SelectedPath = TxtUncPath.Text.Trim();
            if (string.IsNullOrEmpty(SelectedPath) || !SelectedPath.StartsWith(@"\\"))
            {
                MessageBox.Show(this, "Введите корректный UNC-путь (\\\\server\\printer).", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        else
        {
            SelectedMode = DriverFixMode.Auto;
            SelectedPath = "";
        }

        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
