using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;

namespace DocTransform.Models;

public partial class ExcelFileInfo : ObservableObject
{
    [ObservableProperty] private string _fileName;

    [ObservableProperty] private string _filePath;

    [ObservableProperty] private int _rowCount;

    public ExcelFileInfo(string path)
    {
        FilePath = path;
        FileName = Path.GetFileName(path);
    }
}