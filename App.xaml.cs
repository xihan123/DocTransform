using System.Windows;
using DocTransform.Models;
using DocTransform.ViewModels;
using Application = System.Windows.Application;

namespace DocTransform;

/// <summary>
///     Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private MainViewModel _mainViewModel;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // 初始化应用程序设置
        var settings = AppSettings.Instance;
    }

    protected override void OnExit(ExitEventArgs e)
    {
        // 确保在退出时保存设置
        AppSettings.Instance.Save();

        // 调用ViewModel的Dispose方法（如果MainWindow.DataContext是MainViewModel）
        if (MainWindow?.DataContext is MainViewModel viewModel) viewModel.Dispose();

        base.OnExit(e);
    }
}