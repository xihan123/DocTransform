using System.Windows;
using System.Windows.Input;
using DocTransform.ViewModels;
using MaterialDesignThemes.Wpf;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;

namespace DocTransform;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();

        // 添加窗口状态变化的事件处理，更新最大化/还原按钮图标
        StateChanged += MainWindow_StateChanged;

        // 处理窗口最大化时的边距
        SizeChanged += (s, e) =>
        {
            if (WindowState == WindowState.Maximized)
            {
                // 当窗口最大化时，添加边距以避免覆盖任务栏
                var margin = SystemParameters.WindowResizeBorderThickness;
                margin.Top += SystemParameters.CaptionHeight;
                Margin = margin;
            }
            else
            {
                Margin = new Thickness(0);
            }
        };

        // 添加拖放事件处理
        AllowDrop = true;
        Drop += MainWindow_Drop;
        PreviewDragOver += MainWindow_PreviewDragOver;

        // 增强拖放体验
        AllowDrop = true;
        Drop += MainWindow_Drop;
        PreviewDragOver += MainWindow_PreviewDragOver;
        PreviewDragEnter += MainWindow_PreviewDragEnter;
    }

    private void MainWindow_PreviewDragEnter(object sender, DragEventArgs e)
    {
        // 当文件拖入窗口时提供视觉反馈
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            // 在此可以添加高亮效果
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private void MainWindow_PreviewDragOver(object sender, DragEventArgs e)
    {
        // 允许放置文件
        e.Effects = DragDropEffects.Copy;
        e.Handled = true;
    }

    private async void MainWindow_Drop(object sender, DragEventArgs e)
    {
        if (DataContext is MainViewModel viewModel)
            // 调用ViewModel中的处理方法
            await viewModel.HandleFileDrop(e);
    }

    private void MainWindow_StateChanged(object sender, EventArgs e)
    {
        // 更新最大化按钮图标
        if (WindowState == WindowState.Maximized)
        {
            MaximizeButton.ToolTip = "还原";
            ((PackIcon)MaximizeButton.Content).Kind = PackIconKind.WindowRestore;
        }
        else
        {
            MaximizeButton.ToolTip = "最大化";
            ((PackIcon)MaximizeButton.Content).Kind = PackIconKind.WindowMaximize;
        }
    }

    private void ColorZone_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        // 允许拖动窗口
        if (e.ClickCount == 1)
            DragMove();
        else if (e.ClickCount == 2)
            // 双击切换最大化状态
            ToggleMaximize();
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    private void MaximizeButton_Click(object sender, RoutedEventArgs e)
    {
        ToggleMaximize();
    }

    private void ToggleMaximize()
    {
        if (WindowState == WindowState.Maximized)
            WindowState = WindowState.Normal;
        else
            WindowState = WindowState.Maximized;
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}