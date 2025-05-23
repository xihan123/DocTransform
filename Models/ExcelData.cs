namespace DocTransform.Models;

public class ExcelData
{
    public List<string> Headers { get; set; } = new();
    public List<Dictionary<string, string>> Rows { get; set; } = new();
    public HashSet<string> SelectedColumns { get; set; } = new();

    // 添加源文件信息，用于在UI中显示
    public string SourceFileName { get; set; } = string.Empty;
    public int RowCount => Rows.Count;
}