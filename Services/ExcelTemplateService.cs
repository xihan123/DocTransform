using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using DocTransform.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace DocTransform.Services;

/// <summary>
///     Excel模板处理服务，用于处理Excel模板中的占位符替换
/// </summary>
public class ExcelTemplateService
{
    private readonly ImageProcessingService _imageProcessingService;

    public ExcelTemplateService(ImageProcessingService imageProcessingService)
    {
        _imageProcessingService = imageProcessingService;
    }


    /// <summary>
    ///     检查Excel模板文件是否有效
    /// </summary>
    /// <param name="filePath">Excel模板文件路径</param>
    /// <returns>是否为有效的Excel模板</returns>
    public async Task<bool> IsValidTemplateAsync(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return false;

        return await Task.Run(() =>
        {
            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                // 至少要有一个工作表
                return package.Workbook.Worksheets.Count > 0;
            }
            catch
            {
                return false;
            }
        });
    }

    /// <summary>
    ///     从Excel模板中提取所有占位符
    /// </summary>
    /// <param name="filePath">Excel模板文件路径</param>
    /// <returns>占位符集合</returns>
    public async Task<List<string>> ExtractPlaceholdersAsync(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return new List<string>();

        return await Task.Run(() =>
        {
            var placeholders = new HashSet<string>(); // 使用HashSet避免重复
            var regex = new Regex(@"\{([^{}]+)\}");

            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));

                // 遍历所有工作表
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Dimension == null) continue;

                    // 遍历所有单元格
                    for (var row = 1; row <= worksheet.Dimension.End.Row; row++)
                    for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        if (cell?.Value != null)
                        {
                            var cellValue = cell.Value.ToString();
                            var matches = regex.Matches(cellValue);
                            foreach (Match match in matches) placeholders.Add(match.Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"提取Excel占位符时出错: {ex.Message}");
            }

            return new List<string>(placeholders);
        });
    }

    /// <summary>
    ///     处理Excel模板，替换占位符，保留原始格式
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="data">用于替换的数据字典</param>
    /// <param name="progress">进度回调</param>
    /// <returns>处理结果</returns>
    public async Task<(bool Success, string Message)> ProcessTemplateAsync(
        string templatePath,
        string outputPath,
        Dictionary<string, string> data,
        IProgress<int> progress = null)
    {
        if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath)) return (false, "模板文件不存在");

        if (string.IsNullOrEmpty(outputPath)) return (false, "输出路径无效");

        // 创建输出目录
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

        return await Task.Run(() =>
        {
            try
            {
                // 确保输出目录存在
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                // 复制模板文件
                File.Copy(templatePath, outputPath, true);

                using var package = new ExcelPackage(new FileInfo(outputPath));

                var totalSheets = package.Workbook.Worksheets.Count;
                var processedSheets = 0;

                // 遍历所有工作表
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Dimension == null) continue;

                    // 遍历所有单元格
                    for (var row = 1; row <= worksheet.Dimension.End.Row; row++)
                    for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        if (cell?.Value != null && cell.Value is string cellValue)
                        {
                            var newValue = cellValue;
                            var replaced = false;

                            // 替换占位符
                            foreach (var item in data)
                            {
                                var placeholder = $"{{{item.Key}}}";
                                if (newValue.Contains(placeholder))
                                {
                                    // 不直接设置value，而是使用RichText来保留格式
                                    // 先记住原始格式
                                    var originalStyle = cell.Style;
                                    var originalFormat = cell.Style.Numberformat.Format;

                                    // 替换文本
                                    newValue = newValue.Replace(placeholder, item.Value ?? string.Empty);
                                    replaced = true;
                                }
                            }

                            if (replaced)
                            {
                                // 保存原始格式数据
                                var originalStyle = new
                                {
                                    Font = new
                                    {
                                        cell.Style.Font.Name,
                                        cell.Style.Font.Size,
                                        cell.Style.Font.Bold,
                                        cell.Style.Font.Italic,
                                        cell.Style.Font.UnderLine,
                                        cell.Style.Font.Strike,
                                        cell.Style.Font.Color
                                    },
                                    Fill = new
                                    {
                                        cell.Style.Fill.BackgroundColor, cell.Style.Fill.PatternType
                                    },
                                    Border = new
                                    {
                                        cell.Style.Border.Bottom,
                                        cell.Style.Border.Top,
                                        cell.Style.Border.Left,
                                        cell.Style.Border.Right
                                    },
                                    Alignment = new
                                    {
                                        Horizontal = cell.Style.HorizontalAlignment,
                                        Vertical = cell.Style.VerticalAlignment,
                                        cell.Style.WrapText
                                    },
                                    NumberFormat = cell.Style.Numberformat.Format
                                };

                                // 设置新的值
                                cell.Value = newValue;

                                // 确保恢复原始字体设置
                                if (!string.IsNullOrEmpty(originalStyle.Font.Name))
                                    cell.Style.Font.Name = originalStyle.Font.Name;

                                if (originalStyle.Font.Size > 0)
                                    cell.Style.Font.Size = originalStyle.Font.Size;

                                cell.Style.Font.Bold = originalStyle.Font.Bold;
                                cell.Style.Font.Italic = originalStyle.Font.Italic;
                                cell.Style.Font.UnderLine = originalStyle.Font.UnderLine;
                                cell.Style.Font.Strike = originalStyle.Font.Strike;

                                if (originalStyle.Font.Color.Indexed > 0)
                                    cell.Style.Font.Color.SetColor(
                                        ColorTranslator.FromOle(originalStyle.Font.Color.Indexed));
                                else if (!string.IsNullOrEmpty(originalStyle.Font.Color.Rgb))
                                    cell.Style.Font.Color.SetColor(
                                        ColorTranslator.FromHtml($"#{originalStyle.Font.Color.Rgb}"));

                                // 恢复背景色和填充
                                cell.Style.Fill.PatternType = originalStyle.Fill.PatternType;
                                if (originalStyle.Fill.BackgroundColor.Indexed > 0)
                                {
                                    var color = ColorTranslator.FromOle(originalStyle.Fill.BackgroundColor.Indexed);
                                    cell.Style.Fill.BackgroundColor.SetColor(color);
                                }
                                else if (!string.IsNullOrEmpty(originalStyle.Fill.BackgroundColor.Rgb))
                                {
                                    var color = ColorTranslator.FromHtml($"#{originalStyle.Fill.BackgroundColor.Rgb}");
                                    cell.Style.Fill.BackgroundColor.SetColor(color);
                                }

                                // 恢复边框
                                cell.Style.Border.Bottom.Style = originalStyle.Border.Bottom.Style;
                                if (originalStyle.Border.Bottom.Color.Indexed > 0)
                                {
                                    var color = ColorTranslator.FromOle(originalStyle.Border.Bottom.Color.Indexed);
                                    cell.Style.Border.Bottom.Color.SetColor(color);
                                }
                                else if (!string.IsNullOrEmpty(originalStyle.Border.Bottom.Color.Rgb))
                                {
                                    var color = ColorTranslator.FromHtml($"#{originalStyle.Border.Bottom.Color.Rgb}");
                                    cell.Style.Border.Bottom.Color.SetColor(color);
                                }

                                cell.Style.Border.Bottom.Style = originalStyle.Border.Bottom.Style;
                                cell.Style.Border.Top.Style = originalStyle.Border.Top.Style;
                                cell.Style.Border.Left.Style = originalStyle.Border.Left.Style;
                                cell.Style.Border.Right.Style = originalStyle.Border.Right.Style;

                                // 恢复对齐方式
                                cell.Style.HorizontalAlignment = originalStyle.Alignment.Horizontal;
                                cell.Style.VerticalAlignment = originalStyle.Alignment.Vertical;
                                cell.Style.WrapText = originalStyle.Alignment.WrapText;

                                // 恢复数字格式
                                if (!string.IsNullOrEmpty(originalStyle.NumberFormat))
                                    cell.Style.Numberformat.Format = originalStyle.NumberFormat;
                            }
                        }
                    }

                    processedSheets++;
                    progress?.Report(processedSheets * 100 / totalSheets);
                }

                // 保存修改
                package.Save();
                return (true, "Excel模板处理成功");
            }
            catch (Exception ex)
            {
                return (false, $"处理Excel模板时出错: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     处理Excel模板，替换占位符和图片
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="data">用于替换文本占位符的数据字典</param>
    /// <param name="imageDirectories">图片目录列表，用于替换图片占位符</param>
    /// <param name="imageFillMode">图片填充模式</param>
    /// <param name="fillPercentage">填充百分比（0-100）</param>
    /// <param name="progress">进度回调</param>
    /// <returns>处理结果</returns>
    public async Task<(bool Success, string Message)> ProcessTemplateWithImagesAsync(
        string templatePath,
        string outputPath,
        Dictionary<string, string> data,
        IEnumerable<ImageSourceDirectory> imageDirectories,
        ImageFillMode imageFillMode = ImageFillMode.Fit,
        int fillPercentage = 90,
        IProgress<int> progress = null)
    {
        if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath)) return (false, "模板文件不存在");

        if (string.IsNullOrEmpty(outputPath)) return (false, "输出路径无效");

        // 创建输出目录
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

        return await Task.Run(() =>
        {
            try
            {
                // 确保输出目录存在
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                // 复制模板文件
                File.Copy(templatePath, outputPath, true);

                using var package = new ExcelPackage(new FileInfo(outputPath));

                // 预处理图片目录信息，检查有效性
                var validImageDirectories = new List<(ImageSourceDirectory Dir, string MatchValue)>();
                foreach (var imageDir in imageDirectories)
                {
                    if (string.IsNullOrEmpty(imageDir.MatchingColumn) ||
                        !data.TryGetValue(imageDir.MatchingColumn, out var matchValue) ||
                        string.IsNullOrEmpty(matchValue) ||
                        imageDir.ImageFiles.Count == 0)
                        continue;

                    validImageDirectories.Add((imageDir, matchValue));
                }

                var totalSheets = package.Workbook.Worksheets.Count;
                var processedSheets = 0;

                // 遍历所有工作表
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Dimension == null) continue;

                    // 第一轮：处理文本占位符
                    for (var row = 1; row <= worksheet.Dimension.End.Row; row++)
                    for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        if (cell?.Value != null && cell.Value is string cellValue)
                        {
                            var newValue = cellValue;
                            var replaced = false;

                            // 替换文本占位符
                            foreach (var item in data)
                            {
                                var placeholder = $"{{{item.Key}}}";
                                if (newValue.Contains(placeholder))
                                {
                                    newValue = newValue.Replace(placeholder, item.Value ?? string.Empty);
                                    replaced = true;
                                }
                            }

                            if (replaced)
                            {
                                // 保存原始格式数据
                                var originalStyle = new
                                {
                                    Font = new
                                    {
                                        cell.Style.Font.Name,
                                        cell.Style.Font.Size,
                                        cell.Style.Font.Bold,
                                        cell.Style.Font.Italic,
                                        cell.Style.Font.UnderLine,
                                        cell.Style.Font.Strike,
                                        cell.Style.Font.Color
                                    },
                                    Fill = new
                                    {
                                        cell.Style.Fill.BackgroundColor, cell.Style.Fill.PatternType
                                    },
                                    Border = new
                                    {
                                        cell.Style.Border.Bottom,
                                        cell.Style.Border.Top,
                                        cell.Style.Border.Left,
                                        cell.Style.Border.Right
                                    },
                                    Alignment = new
                                    {
                                        Horizontal = cell.Style.HorizontalAlignment,
                                        Vertical = cell.Style.VerticalAlignment,
                                        cell.Style.WrapText
                                    },
                                    NumberFormat = cell.Style.Numberformat.Format
                                };

                                // 设置新的值
                                cell.Value = newValue;

                                // 确保恢复原始字体设置
                                if (!string.IsNullOrEmpty(originalStyle.Font.Name))
                                    cell.Style.Font.Name = originalStyle.Font.Name;

                                if (originalStyle.Font.Size > 0)
                                    cell.Style.Font.Size = originalStyle.Font.Size;

                                cell.Style.Font.Bold = originalStyle.Font.Bold;
                                cell.Style.Font.Italic = originalStyle.Font.Italic;
                                cell.Style.Font.UnderLine = originalStyle.Font.UnderLine;
                                cell.Style.Font.Strike = originalStyle.Font.Strike;

                                if (originalStyle.Font.Color.Indexed > 0)
                                    cell.Style.Font.Color.SetColor(
                                        ColorTranslator.FromOle(originalStyle.Font.Color.Indexed));
                                else if (!string.IsNullOrEmpty(originalStyle.Font.Color.Rgb))
                                    cell.Style.Font.Color.SetColor(
                                        ColorTranslator.FromHtml($"#{originalStyle.Font.Color.Rgb}"));

                                // 恢复背景色和填充
                                cell.Style.Fill.PatternType = originalStyle.Fill.PatternType;
                                if (originalStyle.Fill.BackgroundColor.Indexed > 0)
                                {
                                    var color = ColorTranslator.FromOle(originalStyle.Fill.BackgroundColor.Indexed);
                                    cell.Style.Fill.BackgroundColor.SetColor(color);
                                }
                                else if (!string.IsNullOrEmpty(originalStyle.Fill.BackgroundColor.Rgb))
                                {
                                    var color = ColorTranslator.FromHtml($"#{originalStyle.Fill.BackgroundColor.Rgb}");
                                    cell.Style.Fill.BackgroundColor.SetColor(color);
                                }

                                // 恢复边框
                                cell.Style.Border.Bottom.Style = originalStyle.Border.Bottom.Style;
                                if (originalStyle.Border.Bottom.Color.Indexed > 0)
                                {
                                    var color = ColorTranslator.FromOle(originalStyle.Border.Bottom.Color.Indexed);
                                    cell.Style.Border.Bottom.Color.SetColor(color);
                                }
                                else if (!string.IsNullOrEmpty(originalStyle.Border.Bottom.Color.Rgb))
                                {
                                    var color = ColorTranslator.FromHtml($"#{originalStyle.Border.Bottom.Color.Rgb}");
                                    cell.Style.Border.Bottom.Color.SetColor(color);
                                }

                                cell.Style.Border.Top.Style = originalStyle.Border.Top.Style;
                                cell.Style.Border.Left.Style = originalStyle.Border.Left.Style;
                                cell.Style.Border.Right.Style = originalStyle.Border.Right.Style;

                                // 恢复对齐方式
                                cell.Style.HorizontalAlignment = originalStyle.Alignment.Horizontal;
                                cell.Style.VerticalAlignment = originalStyle.Alignment.Vertical;
                                cell.Style.WrapText = originalStyle.Alignment.WrapText;

                                // 恢复数字格式
                                if (!string.IsNullOrEmpty(originalStyle.NumberFormat))
                                    cell.Style.Numberformat.Format = originalStyle.NumberFormat;
                            }
                        }
                    }

                    // 第二轮：处理图片占位符
                    // 图片处理不修改文本格式，保持不变...
                    for (var row = 1; row <= worksheet.Dimension.End.Row; row++)
                    for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        if (cell?.Value != null && cell.Value is string cellValue)
                            // 检查每个图片目录的占位符
                            foreach (var (imageDir, matchValue) in validImageDirectories)
                            {
                                var placeholder = imageDir.PlaceholderName;
                                if (cellValue.Contains(placeholder))
                                {
                                    // 查找匹配的图片
                                    var matchingImagePath = _imageProcessingService.FindMatchingImage(
                                        imageDir.ImageFiles, matchValue);

                                    if (!string.IsNullOrEmpty(matchingImagePath))
                                    {
                                        // 清空单元格值
                                        cell.Value = null;

                                        // 插入图片
                                        InsertImageIntoCell(worksheet, row, col, matchingImagePath, imageFillMode,
                                            fillPercentage);
                                    }
                                }
                            }
                    }

                    processedSheets++;
                    progress?.Report(processedSheets * 100 / totalSheets);
                }

                // 保存修改
                package.Save();
                return (true, "Excel模板处理成功");
            }
            catch (Exception ex)
            {
                return (false, $"处理Excel模板时出错: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     在Excel单元格中插入图片，精确缩放以匹配单元格大小并完美嵌入
    /// </summary>
    /// <param name="worksheet">工作表</param>
    /// <param name="row">行索引</param>
    /// <param name="col">列索引</param>
    /// <param name="imagePath">图片路径</param>
    /// <param name="fillMode">图片填充模式</param>
    /// <param name="fillPercentage">填充百分比（0-100）</param>
    private void InsertImageIntoCell(
        ExcelWorksheet worksheet,
        int row,
        int col,
        string imagePath,
        ImageFillMode fillMode = ImageFillMode.Fit,
        int fillPercentage = 90)
    {
        try
        {
            // 保存单元格格式，以便在操作后恢复
            var cell = worksheet.Cells[row, col];

            // 记录单元格样式信息，以备需要
            var cellStyle = cell.Style;

            using (var image = _imageProcessingService.LoadImageFromFile(imagePath))
            {
                if (image == null) return;

                // 为图片创建唯一名称
                var picName = $"Image_R{row}C{col}_{Guid.NewGuid().ToString("N").Substring(0, 8)}";
                var picture = worksheet.Drawings.AddPicture(picName, imagePath);

                // 更精确地计算单元格宽度（像素）
                // EPPlus中列宽单位是字符数，1字符约等于8像素（标准字体下）
                var columnWidthInChars = worksheet.Column(col).Width;
                if (columnWidthInChars <= 0) columnWidthInChars = 10;
                var cellWidthPx = columnWidthInChars * 8;

                // 计算行高（像素）
                // EPPlus中行高单位是点，1点约等于1.33像素
                var rowHeightInPts = worksheet.Row(row).Height;
                if (rowHeightInPts <= 0) rowHeightInPts = 15;
                var cellHeightPx = rowHeightInPts * 1.33;

                // 应用填充百分比
                var fillRatio = Math.Max(10, Math.Min(100, fillPercentage)) / 100.0;
                var effectiveWidth = cellWidthPx * fillRatio;
                var effectiveHeight = cellHeightPx * fillRatio;

                // 计算图片缩放比例
                var scaleX = effectiveWidth / image.Width;
                var scaleY = effectiveHeight / image.Height;

                // 根据填充模式选择缩放比例
                double scale;
                switch (fillMode)
                {
                    case ImageFillMode.Fill:
                        // 填充模式：使用较大的缩放比例，使图片填满单元格的一个维度
                        scale = Math.Max(scaleX, scaleY);
                        break;
                    case ImageFillMode.Stretch:
                        // 拉伸模式：独立缩放宽度和高度，完全填充单元格
                        scale = 1.0; // 临时值，不会使用
                        break;
                    case ImageFillMode.Fit:
                    default:
                        // 适应模式：使用较小的缩放比例，确保图片完全适应单元格
                        scale = Math.Min(scaleX, scaleY);
                        break;
                }

                // 计算缩放后的图片尺寸
                double scaledWidth, scaledHeight;

                if (fillMode == ImageFillMode.Stretch)
                {
                    // 拉伸模式下，宽高独立缩放
                    scaledWidth = effectiveWidth;
                    scaledHeight = effectiveHeight;
                }
                else
                {
                    // 其他模式下，按比例缩放
                    scaledWidth = image.Width * scale;
                    scaledHeight = image.Height * scale;
                }

                // 计算图片在单元格中的位置（居中）
                var offsetX = (cellWidthPx - scaledWidth) / 2;
                var offsetY = (cellHeightPx - scaledHeight) / 2;

                // 确保偏移量不为负
                offsetX = Math.Max(0, offsetX);
                offsetY = Math.Max(0, offsetY);

                // 设置图片位置
                picture.SetPosition(row - 1, (int)offsetY, col - 1, (int)offsetX);

                // 精确设置图片大小（像素）
                picture.SetSize((int)scaledWidth, (int)scaledHeight);

                // 设置图片其他属性，以确保其在单元格中显示正确
                picture.EditAs = eEditAs.Absolute; // 使用绝对位置
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"插入图片时出错: {ex.Message}");
        }
    }
}