using System.Diagnostics;
using System.IO;
using DocTransform.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocTransform.Services;

/// <summary>
///     Word文档处理服务
/// </summary>
public class WordService
{
    /// <summary>
    ///     异步处理Word模板并替换占位符
    /// </summary>
    /// <param name="templatePath">Word模板路径</param>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="data">行数据</param>
    /// <param name="progress">进度回调</param>
    /// <returns>处理结果</returns>
    public async Task<ProcessingResult> ProcessTemplateAsync(
        string templatePath,
        string outputPath,
        Dictionary<string, string> data,
        IProgress<int>? progress = null)
    {
        return await Task.Run(() =>
        {
            try
            {
                // 验证文件存在
                if (!File.Exists(templatePath)) return ProcessingResult.Fail("Word模板文件不存在");

                // 确保输出目录存在
                var outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // 复制模板文件到输出路径
                File.Copy(templatePath, outputPath, true);

                using var document = WordprocessingDocument.Open(outputPath, true);
                var body = document.MainDocumentPart?.Document.Body;

                if (body == null) return ProcessingResult.Fail("无法读取Word文档内容");

                // 替换正文中的占位符
                ReplaceTextInElement(body, data);

                // 替换页眉中的占位符
                if (document.MainDocumentPart?.HeaderParts != null)
                    foreach (var headerPart in document.MainDocumentPart.HeaderParts)
                        ReplaceTextInElement(headerPart.Header, data);

                // 替换页脚中的占位符
                if (document.MainDocumentPart?.FooterParts != null)
                    foreach (var footerPart in document.MainDocumentPart.FooterParts)
                        ReplaceTextInElement(footerPart.Footer, data);

                document.Save();
                progress?.Report(100);

                return ProcessingResult.Succeed("文档生成成功", outputPath);
            }
            catch (Exception ex)
            {
                return ProcessingResult.Fail($"处理模板时出错: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     替换OpenXML元素中的所有文本占位符
    /// </summary>
    /// <param name="element">OpenXML元素</param>
    /// <param name="data">替换数据</param>
    private void ReplaceTextInElement(OpenXmlElement element, Dictionary<string, string> data)
    {
        try
        {
            // 获取所有段落
            var paragraphs = element.Descendants<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                var paragraphText = paragraph.InnerText;
                var paragraphModified = false;

                foreach (var kvp in data)
                {
                    var placeholder = $"{{{kvp.Key}}}";
                    if (paragraphText.Contains(placeholder, StringComparison.OrdinalIgnoreCase))
                    {
                        // 需要替换这个段落中的文本
                        paragraphModified = true;

                        // 获取段落中的所有文本运行
                        var runs = paragraph.Elements<Run>().ToList();

                        // 用于追踪我们需要创建的运行
                        var newRuns = new List<Run>();

                        // 提取当前段落的所有文本
                        var fullText = string.Empty;
                        foreach (var run in runs) fullText += run.InnerText;

                        // 替换所有占位符
                        foreach (var item in data)
                        {
                            var itemPlaceholder = $"{{{item.Key}}}";
                            fullText = fullText.Replace(itemPlaceholder, item.Value ?? string.Empty,
                                StringComparison.OrdinalIgnoreCase);
                        }

                        // 清空段落中的文本运行
                        foreach (var run in runs) run.Remove();

                        // 创建新的文本运行并添加到段落
                        var newRun = new Run();
                        var text = new Text(fullText);
                        newRun.AppendChild(text);
                        paragraph.AppendChild(newRun);

                        break;
                    }
                }

                // 如果段落没有修改，也尝试替换单个文本运行中的占位符
                if (!paragraphModified)
                {
                    var textElements = paragraph.Descendants<Text>().ToList();

                    foreach (var text in textElements)
                    {
                        var originalText = text.Text ?? string.Empty;
                        var newText = originalText;

                        // 查找并替换所有匹配的占位符
                        foreach (var kvp in data)
                        {
                            var placeholder = $"{{{kvp.Key}}}";
                            if (newText.Contains(placeholder, StringComparison.OrdinalIgnoreCase))
                                newText = newText.Replace(placeholder, kvp.Value ?? string.Empty,
                                    StringComparison.OrdinalIgnoreCase);
                        }

                        // 只在文本变化时更新
                        if (newText != originalText) text.Text = newText;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // 记录异常，但不中断处理
            Debug.WriteLine($"替换文本时出错: {ex.Message}");
        }
    }

    /// <summary>
    ///     检查Word模板是否有效
    /// </summary>
    public async Task<bool> IsValidTemplateAsync(string templatePath)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (!File.Exists(templatePath))
                    return false;

                using var document = WordprocessingDocument.Open(templatePath, false);
                return document.MainDocumentPart?.Document.Body != null;
            }
            catch
            {
                return false;
            }
        });
    }
}