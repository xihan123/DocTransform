using System.Diagnostics;
using System.IO;

namespace DocTransform.Services;

/// <summary>
///     图片处理服务，提供图片文件扫描、匹配和处理功能
/// </summary>
public class ImageProcessingService
{
    /// <summary>
    ///     获取支持的图片文件扩展名列表
    /// </summary>
    public IEnumerable<string> SupportedImageExtensions => new[]
    {
        ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp", ".ico"
    };

    /// <summary>
    ///     扫描指定目录，获取所有支持的图片文件
    /// </summary>
    /// <param name="directoryPath">目录路径</param>
    /// <returns>图片文件路径列表</returns>
    public async Task<List<string>> ScanDirectoryForImagesAsync(string directoryPath)
    {
        if (string.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath)) return new List<string>();

        return await Task.Run(() =>
        {
            try
            {
                return Directory.GetFiles(directoryPath)
                    .Where(file => SupportedImageExtensions.Contains(
                        Path.GetExtension(file).ToLowerInvariant()))
                    .ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"扫描图片目录时出错: {ex.Message}");
                return new List<string>();
            }
        });
    }

    /// <summary>
    ///     根据匹配值查找对应的图片文件
    /// </summary>
    /// <param name="imageFiles">图片文件列表</param>
    /// <param name="matchValue">匹配值（通常是列值）</param>
    /// <returns>匹配的图片文件路径，如果没有匹配则返回null</returns>
    public string FindMatchingImage(List<string> imageFiles, string matchValue)
    {
        if (string.IsNullOrEmpty(matchValue) || imageFiles == null || imageFiles.Count == 0) return null;

        // 移除扩展名后进行比较
        var valueWithoutExtension = Path.GetFileNameWithoutExtension(matchValue);

        // 查找完全匹配的文件名
        var exactMatch = imageFiles.FirstOrDefault(img =>
            Path.GetFileNameWithoutExtension(img).Equals(
                valueWithoutExtension,
                StringComparison.OrdinalIgnoreCase));

        if (!string.IsNullOrEmpty(exactMatch)) return exactMatch;

        // 如果没有完全匹配，查找包含匹配值的文件名
        return imageFiles.FirstOrDefault(img =>
            Path.GetFileNameWithoutExtension(img).Contains(
                valueWithoutExtension,
                StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    ///     从文件加载图片
    /// </summary>
    /// <param name="imagePath">图片文件路径</param>
    /// <returns>图片对象，如果加载失败则返回null</returns>
    public Image LoadImageFromFile(string imagePath)
    {
        if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath)) return null;

        try
        {
            return Image.FromFile(imagePath);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"加载图片失败: {ex.Message}");
            return null;
        }
    }
}