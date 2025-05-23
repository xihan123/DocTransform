namespace DocTransform.Models;

/// <summary>
///     图片填充模式枚举
/// </summary>
public enum ImageFillMode
{
    /// <summary>
    ///     保持图片比例，适应单元格（默认）
    /// </summary>
    Fit = 0,

    /// <summary>
    ///     保持图片比例，最大化填充单元格
    /// </summary>
    Fill = 1,

    /// <summary>
    ///     拉伸以完全填充单元格（可能变形）
    /// </summary>
    Stretch = 2
}

/// <summary>
///     枚举显示辅助类，用于在UI中显示枚举值
/// </summary>
public class ImageFillModeItem
{
    public ImageFillMode Value { get; set; }
    public string DisplayName { get; set; }

    public static List<ImageFillModeItem> GetAll()
    {
        return new List<ImageFillModeItem>
        {
            new() { Value = ImageFillMode.Fit, DisplayName = "适应 - 确保整张图片可见" },
            new() { Value = ImageFillMode.Fill, DisplayName = "填充 - 最大化填充单元格" },
            new() { Value = ImageFillMode.Stretch, DisplayName = "拉伸 - 完全填充单元格" }
        };
    }
}