using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace DocTransform.Services;

/// <summary>
/// Excel模板处理服务，用于处理Excel模板中的占位符替换
/// </summary>
public class ExcelTemplateService
{
    /// <summary>
    /// 检查Excel模板文件是否有效
    /// </summary>
    /// <param name="filePath">Excel模板文件路径</param>
    /// <returns>是否为有效的Excel模板</returns>
    public async Task<bool> IsValidTemplateAsync(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
        {
            return false;
        }

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
    /// 从Excel模板中提取所有占位符
    /// </summary>
    /// <param name="filePath">Excel模板文件路径</param>
    /// <returns>占位符集合</returns>
    public async Task<List<string>> ExtractPlaceholdersAsync(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
        {
            return new List<string>();
        }

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
                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var cell = worksheet.Cells[row, col];
                            if (cell?.Value != null)
                            {
                                string cellValue = cell.Value.ToString();
                                var matches = regex.Matches(cellValue);
                                foreach (Match match in matches)
                                {
                                    placeholders.Add(match.Value);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"提取Excel占位符时出错: {ex.Message}");
            }

            return new List<string>(placeholders);
        });
    }

    /// <summary>
    /// 处理Excel模板，替换占位符
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
        if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
        {
            return (false, "模板文件不存在");
        }

        if (string.IsNullOrEmpty(outputPath))
        {
            return (false, "输出路径无效");
        }

        // 创建输出目录
        string outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        return await Task.Run(() =>
        {
            try
            {
                // 确保输出目录存在
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                // 复制模板文件
                File.Copy(templatePath, outputPath, true);

                using var package = new ExcelPackage(new FileInfo(outputPath));

                int totalSheets = package.Workbook.Worksheets.Count;
                int processedSheets = 0;

                // 遍历所有工作表
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Dimension == null) continue;

                    // 遍历所有单元格
                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var cell = worksheet.Cells[row, col];
                            if (cell?.Value != null && cell.Value is string cellValue)
                            {
                                string newValue = cellValue;
                                bool replaced = false;

                                // 替换占位符
                                foreach (var item in data)
                                {
                                    string placeholder = $"{{{item.Key}}}";
                                    if (newValue.Contains(placeholder))
                                    {
                                        newValue = newValue.Replace(placeholder, item.Value ?? string.Empty);
                                        replaced = true;
                                    }
                                }

                                if (replaced)
                                {
                                    cell.Value = newValue;
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
}