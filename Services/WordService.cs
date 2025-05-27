using System.Diagnostics;
using System.IO;
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
    ///     处理Word模板，替换占位符，保留原始格式
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
                // 复制模板文件
                File.Copy(templatePath, outputPath, true);

                // 打开Word文档
                using (var document = WordprocessingDocument.Open(outputPath, true))
                {
                    // 获取文档主体
                    var body = document.MainDocumentPart.Document.Body;
                    var totalTextElements = 0;

                    // 获取所有文本元素以计算进度
                    CountTextElements(body, ref totalTextElements);
                    var processedElements = 0;

                    // 替换正文中的占位符
                    ReplacePlaceholdersInElement(body, data, ref processedElements, totalTextElements, progress);

                    // 处理页眉和页脚
                    if (document.MainDocumentPart.HeaderParts != null)
                        foreach (var headerPart in document.MainDocumentPart.HeaderParts)
                            ReplacePlaceholdersInElement(headerPart.Header, data, ref processedElements,
                                totalTextElements, progress);

                    if (document.MainDocumentPart.FooterParts != null)
                        foreach (var footerPart in document.MainDocumentPart.FooterParts)
                            ReplacePlaceholdersInElement(footerPart.Footer, data, ref processedElements,
                                totalTextElements, progress);

                    // 保存更改
                    document.MainDocumentPart.Document.Save();

                    // 最终进度报告
                    progress?.Report(100);
                }

                return (true, "Word模板处理成功");
            }
            catch (Exception ex)
            {
                return (false, $"处理Word模板时出错: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     计算OpenXML元素中的文本元素数量
    /// </summary>
    /// <param name="element">要计数的元素</param>
    /// <param name="count">计数结果</param>
    private void CountTextElements(OpenXmlElement element, ref int count)
    {
        if (element == null) return;

        // 计数文本元素
        if (element is Text || element is SimpleField) count++;

        // 递归处理子元素
        foreach (var child in element.Elements()) CountTextElements(child, ref count);
    }

    /// <summary>
    ///     替换OpenXML元素中的占位符（保留格式）
    /// </summary>
    /// <param name="element">OpenXML元素</param>
    /// <param name="data">替换数据</param>
    /// <param name="processedElements">已处理元素计数</param>
    /// <param name="totalElements">总元素数量</param>
    /// <param name="progress">进度回调</param>
    private void ReplacePlaceholdersInElement(
        OpenXmlElement element,
        Dictionary<string, string> data,
        ref int processedElements,
        int totalElements,
        IProgress<int> progress)
    {
        if (element == null) return;

        // 处理段落和表格等复杂元素
        if (element is Paragraph paragraph)
        {
            // 处理段落中的每个Run
            ReplacePlaceholdersInParagraph(paragraph, data, ref processedElements, totalElements, progress);
        }
        else if (element is Text text)
        {
            // 处理单独的文本元素
            var originalText = text.Text;
            var newText = originalText;
            var replaced = false;

            foreach (var item in data)
            {
                var placeholder = $"{{{item.Key}}}";
                if (newText.Contains(placeholder))
                {
                    // 替换文本，但保持原有格式
                    newText = newText.Replace(placeholder, item.Value ?? string.Empty);
                    replaced = true;
                }
            }

            if (replaced) text.Text = newText;

            processedElements++;
            if (totalElements > 0 && progress != null) progress.Report(processedElements * 100 / totalElements);
        }
        // 处理其他常见包含文本的元素类型
        else if (element is SimpleField simpleField)
        {
            var originalText = simpleField.InnerText; // 使用 InnerText 获取文本内容
            var newText = originalText;
            var replaced = false;

            foreach (var item in data)
            {
                var placeholder = $"{{{item.Key}}}";
                if (newText.Contains(placeholder))
                {
                    // 替换文本
                    newText = newText.Replace(placeholder, item.Value ?? string.Empty);
                    replaced = true;
                }
            }

            if (replaced)
            {
                // SimpleField 没有直接的 Text 属性，因此需要通过其他方式更新内容
                simpleField.RemoveAllChildren(); // 清空现有子元素
                simpleField.AppendChild(new Text(newText)); // 添加新的文本子元素
            }

            processedElements++;
            if (totalElements > 0 && progress != null) progress.Report(processedElements * 100 / totalElements);
        }

        // 递归处理所有子元素
        var childElements = element.Elements().ToList();
        foreach (var child in childElements)
            ReplacePlaceholdersInElement(child, data, ref processedElements, totalElements, progress);
    }

    /// <summary>
    ///     替换段落中的占位符，保留格式
    /// </summary>
    /// <param name="paragraph">段落元素</param>
    /// <param name="data">替换数据</param>
    /// <param name="processedElements">已处理元素计数</param>
    /// <param name="totalElements">总元素数量</param>
    /// <param name="progress">进度回调</param>
    private void ReplacePlaceholdersInParagraph(
        Paragraph paragraph,
        Dictionary<string, string> data,
        ref int processedElements,
        int totalElements,
        IProgress<int> progress)
    {
        if (paragraph == null) return;

        // 特殊情况：跨多个Run的占位符处理

        // 1. 收集段落中的所有文本
        var fullText = "";
        var runPositions = new Dictionary<int, Run>();
        var currentPos = 0;

        foreach (var run in paragraph.Elements<Run>())
        foreach (var textElement in run.Elements<Text>())
        {
            var runText = textElement.Text;
            fullText += runText;
            runPositions[currentPos] = run;
            currentPos += runText.Length;
        }

        // 2. 检查是否有占位符跨越多个Run
        var hasReplacement = false;
        foreach (var item in data)
        {
            var placeholder = $"{{{item.Key}}}";
            if (fullText.Contains(placeholder))
            {
                hasReplacement = true;
                break;
            }
        }

        // 3. 如果有跨Run的占位符，进行特殊处理
        if (hasReplacement)
        {
            // 方法1：从段落的源文本开始重建段落，保留格式
            var combinedText = fullText;

            // 替换完整文本中的所有占位符
            foreach (var item in data)
            {
                var placeholder = $"{{{item.Key}}}";
                combinedText = combinedText.Replace(placeholder, item.Value ?? string.Empty);
            }

            // 如果有任何改变，则在第一个Run中包含所有文本，并删除其他Run
            if (combinedText != fullText)
            {
                var runs = paragraph.Elements<Run>().ToList();
                if (runs.Any())
                {
                    // 保存第一个Run并保留其属性（格式）
                    var firstRun = runs.First();

                    // 清除该Run中的所有文本
                    var textsInFirstRun = firstRun.Elements<Text>().ToList();
                    foreach (var text in textsInFirstRun) text.Remove();

                    // 添加新文本，保留格式
                    firstRun.AppendChild(new Text(combinedText));

                    // 删除其他Run以避免重复
                    for (var i = 1; i < runs.Count; i++) runs[i].Remove();
                }
            }
        }
        else
        {
            // 如果没有跨Run占位符，则正常处理每个Run
            var runs = paragraph.Elements<Run>().ToList();
            foreach (var run in runs)
            {
                var texts = run.Elements<Text>().ToList();
                foreach (var text in texts)
                    ReplacePlaceholdersInElement(text, data, ref processedElements, totalElements, progress);
            }
        }
    }

    /// <summary>
    ///     从Word模板中提取所有占位符
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <returns>占位符列表</returns>
    public async Task<List<string>> ExtractPlaceholdersAsync(string templatePath)
    {
        if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath)) return new List<string>();

        return await Task.Run(() =>
        {
            var placeholders = new HashSet<string>();
            try
            {
                using (var document = WordprocessingDocument.Open(templatePath, false))
                {
                    // 从主文档部分提取占位符
                    ExtractPlaceholdersFromElement(document.MainDocumentPart.Document.Body, placeholders);

                    // 从页眉页脚提取占位符
                    if (document.MainDocumentPart.HeaderParts != null)
                        foreach (var headerPart in document.MainDocumentPart.HeaderParts)
                            ExtractPlaceholdersFromElement(headerPart.Header, placeholders);

                    if (document.MainDocumentPart.FooterParts != null)
                        foreach (var footerPart in document.MainDocumentPart.FooterParts)
                            ExtractPlaceholdersFromElement(footerPart.Footer, placeholders);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"提取Word占位符时出错: {ex.Message}");
            }

            return new List<string>(placeholders);
        });
    }

    /// <summary>
    ///     从OpenXML元素中提取占位符
    /// </summary>
    /// <param name="element">OpenXML元素</param>
    /// <param name="placeholders">占位符集合</param>
    private void ExtractPlaceholdersFromElement(OpenXmlElement element, HashSet<string> placeholders)
    {
        if (element == null) return;

        if (element is Text text)
            ExtractPlaceholdersFromText(text.Text, placeholders);
        else if (element is SimpleField simpleField) ExtractPlaceholdersFromText(simpleField.InnerText, placeholders);

        // 递归处理子元素
        foreach (var child in element.Elements()) ExtractPlaceholdersFromElement(child, placeholders);
    }

    /// <summary>
    ///     从文本中提取占位符
    /// </summary>
    /// <param name="text">文本内容</param>
    /// <param name="placeholders">占位符集合</param>
    private void ExtractPlaceholdersFromText(string text, HashSet<string> placeholders)
    {
        if (string.IsNullOrEmpty(text)) return;

        var startIndex = 0;
        while ((startIndex = text.IndexOf('{', startIndex)) >= 0)
        {
            var endIndex = text.IndexOf('}', startIndex);
            if (endIndex > startIndex)
            {
                var placeholder = text.Substring(startIndex, endIndex - startIndex + 1);
                placeholders.Add(placeholder);
                startIndex = endIndex + 1;
            }
            else
            {
                // 没有找到匹配的右花括号，结束搜索
                break;
            }
        }
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