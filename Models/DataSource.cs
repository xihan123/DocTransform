using CommunityToolkit.Mvvm.ComponentModel;

namespace DocTransform.Models;

public class DataSource : ObservableObject
{
    private string _displayName;

    private string _filePath;

    private bool _isEnabled = true;

    private int _priority;
    public string Id { get; } = Guid.NewGuid().ToString();

    public string FilePath
    {
        get => _filePath;
        set => SetProperty(ref _filePath, value);
    }

    public string DisplayName
    {
        get => _displayName;
        set => SetProperty(ref _displayName, value);
    }

    public int Priority
    {
        get => _priority;
        set => SetProperty(ref _priority, value);
    }

    public bool IsEnabled
    {
        get => _isEnabled;
        set => SetProperty(ref _isEnabled, value);
    }

    // 数据源中的列
    public List<string> AvailableColumns { get; set; } = new();

    // 行数据
    public List<Dictionary<string, string>> Rows { get; set; } = new();

    // 当前数据源添加的时间（用于排序）
    public DateTime AddedTime { get; } = DateTime.Now;
}