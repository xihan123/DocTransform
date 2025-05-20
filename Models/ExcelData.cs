namespace DocTransform.Models;

/// <summary>
///     表示从Excel文件中提取的数据
/// </summary>
public class ExcelData
{
    /// <summary>
    ///     Excel文件中的列标题
    /// </summary>
    public List<string> Headers { get; set; } = [];

    /// <summary>
    ///     Excel文件中的数据行
    /// </summary>
    public List<Dictionary<string, string>> Rows { get; set; } = [];

    /// <summary>
    ///     用户选择的列
    /// </summary>
    public HashSet<string> SelectedColumns { get; set; } = [];
}