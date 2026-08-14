using System.Windows;
using Microsoft.Win32;
using WFix.App.ViewModels;

namespace WFix.App;

public partial class RemoteCenterWindow : Window
{
    private readonly RemoteCenterViewModel _viewModel;

    public RemoteCenterWindow(RemoteCenterViewModel viewModel)
    {
        _viewModel = viewModel;
        InitializeComponent();
        DataContext = viewModel;
    }

    private async void SaveCredential_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await _viewModel.SaveCredentialAsync(CredentialPassword.Password);
            CredentialPassword.Clear();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Credential Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void ExportSupportBundle_Click(object sender, RoutedEventArgs e)
    {
        if (_viewModel.Runs.Count == 0)
        {
            MessageBox.Show("Сначала выполните пакет или диагностику.", "Support bundle", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new SaveFileDialog
        {
            Title = "Экспорт W-Fix support bundle",
            Filter = "ZIP archive (*.zip)|*.zip",
            FileName = $"W-Fix-support-{DateTime.Now:yyyyMMdd-HHmmss}.zip"
        };
        if (dialog.ShowDialog(this) != true)
            return;

        try
        {
            var path = await _viewModel.ExportSupportBundleAsync(dialog.FileName);
            MessageBox.Show($"Обезличенный архив создан:\n{path}", "Support bundle", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Support bundle", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        _viewModel.Dispose();
        base.OnClosed(e);
    }
}
