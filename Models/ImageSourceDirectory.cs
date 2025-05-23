namespace DocTransform.Models;

/// <summary>
///     图片源目录，用于存储图片目录信息和对应的匹配设置
/// </summary>
public class ImageSourceDirectory
{
    /// <summary>
    ///     目录完整路径
    /// </summary>
    public string DirectoryPath { get; set; } = string.Empty;

    /// <summary>
    ///     目录名称，将用作占位符名称
    /// </summary>
    public string DirectoryName { get; set; } = string.Empty;

    /// <summary>
    ///     用于匹配图片的数据列名
    /// </summary>
    public string MatchingColumn { get; set; } = string.Empty;

    /// <summary>
    ///     图片文件列表
    /// </summary>
    public List<string> ImageFiles { get; set; } = new();

    /// <summary>
    ///     占位符名称，格式为 {目录名}
    /// </summary>
    public string PlaceholderName => $"{{{DirectoryName}}}";

    /// <summary>
    ///     图片文件数量
    /// </summary>
    public int ImageCount => ImageFiles.Count;
}