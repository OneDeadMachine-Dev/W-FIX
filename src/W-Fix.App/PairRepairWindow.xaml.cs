using System.Windows;
using Microsoft.Win32;
using WFix.App.ViewModels;

namespace WFix.App;

public partial class PairRepairWindow : Window
{
    private readonly PairRepairViewModel _viewModel;

    public PairRepairWindow(PairRepairViewModel viewModel)
    {
        _viewModel = viewModel;
        InitializeComponent();
        DataContext = viewModel;
    }

    private async void ExportInvitation_Click(object sender, RoutedEventArgs e)
    {
        var dialog = SaveDialog("W-Fix Pair invitation (*.wfixpair)|*.wfixpair", $"{Environment.MachineName}-{DateTime.Now:HHmm}.wfixpair");
        if (dialog.ShowDialog(this) == true)
            await ShowErrorsAsync(() => _viewModel.ExportInvitationAsync(dialog.FileName));
    }

    private async void ImportInvitation_Click(object sender, RoutedEventArgs e)
    {
        var dialog = OpenDialog("W-Fix Pair invitation (*.wfixpair)|*.wfixpair");
        if (dialog.ShowDialog(this) == true)
            await ShowErrorsAsync(() => _viewModel.ImportInvitationAsync(dialog.FileName));
    }

    private async void ExportOffline_Click(object sender, RoutedEventArgs e)
    {
        var dialog = SaveDialog("W-Fix signed snapshot (*.wfixpair)|*.wfixpair", $"{Environment.MachineName}-snapshot.wfixpair");
        if (dialog.ShowDialog(this) == true)
            await ShowErrorsAsync(() => _viewModel.ExportOfflineSnapshotAsync(dialog.FileName));
    }

    private async void ImportOffline_Click(object sender, RoutedEventArgs e)
    {
        var dialog = OpenDialog("W-Fix signed snapshot (*.wfixpair)|*.wfixpair");
        if (dialog.ShowDialog(this) == true)
            await ShowErrorsAsync(() => _viewModel.ImportOfflineSnapshotAsync(dialog.FileName));
    }

    private async void SaveNetworkCredential_Click(object sender, RoutedEventArgs e)
    {
        var password = NetworkPassword.Password;
        NetworkPassword.Clear();
        await ShowErrorsAsync(() => _viewModel.SaveNetworkCredentialAsync(password));
    }

    private async Task ShowErrorsAsync(Func<Task> operation)
    {
        try { await operation(); }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Pair Repair", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    private static SaveFileDialog SaveDialog(string filter, string name) => new() { Filter = filter, FileName = name, AddExtension = true };
    private static OpenFileDialog OpenDialog(string filter) => new() { Filter = filter, CheckFileExists = true };

    protected override async void OnClosed(EventArgs e)
    {
        await _viewModel.DisposeAsync();
        base.OnClosed(e);
    }
}
