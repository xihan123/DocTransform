using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace DocTransform.Services;

/// <summary>
///     Word文档处理服务
/// </summary>
public class WordService
{
    /// <summary>
    /// 处理Word模板，替换占位符，保留原始格式
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

        string outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        return await Task.Run(() =>
        {
            try
            {
                // 复制模板文件
                File.Copy(templatePath, outputPath, true);

                using (WordprocessingDocument document = WordprocessingDocument.Open(outputPath, true))
                {
                    // 获取所有需要替换的内容
                    var placeholderDict = CreatePlaceholderDict(data);

                    // 处理主文档部分
                    var mainPart = document.MainDocumentPart;
                    int totalElements = CountParagraphs(mainPart);
                    int processedElements = 0;

                    // 处理文档正文
                    if (mainPart?.Document?.Body != null)
                    {
                        ProcessDocumentContent(mainPart.Document.Body, placeholderDict, ref processedElements,
                            totalElements, progress);
                    }

                    // 处理页眉和页脚
                    if (mainPart != null)
                    {
                        foreach (var headerPart in mainPart.HeaderParts)
                        {
                            ProcessDocumentContent(headerPart.Header, placeholderDict, ref processedElements,
                                totalElements, progress);
                        }

                        foreach (var footerPart in mainPart.FooterParts)
                        {
                            ProcessDocumentContent(footerPart.Footer, placeholderDict, ref processedElements,
                                totalElements, progress);
                        }
                    }

                    // 保存文档
                    document.Save();
                }

                progress?.Report(100);
                return (true, "Word模板处理成功");
            }
            catch (Exception ex)
            {
                return (false, $"处理Word模板时出错: {ex.Message}");
            }
        });
    }

    /// <summary>
    /// 创建替换占位符的字典
    /// </summary>
    private Dictionary<string, string> CreatePlaceholderDict(Dictionary<string, string> data)
    {
        var result = new Dictionary<string, string>();
        foreach (var item in data)
        {
            string placeholder = $"{{{item.Key}}}";
            result[placeholder] = item.Value ?? string.Empty;
        }

        return result;
    }

    /// <summary>
    /// 统计文档中的段落数量，用于计算进度
    /// </summary>
    private int CountParagraphs(MainDocumentPart mainPart)
    {
        if (mainPart == null) return 1;

        int count = 0;

        // 计算主体中的段落
        if (mainPart.Document?.Body != null)
        {
            count += mainPart.Document.Body.Descendants<Paragraph>().Count();
            count += mainPart.Document.Body.Descendants<Table>()
                .SelectMany(t => t.Descendants<TableCell>())
                .SelectMany(c => c.Descendants<Paragraph>())
                .Count();
        }

        // 计算页眉中的段落
        foreach (var headerPart in mainPart.HeaderParts)
        {
            count += headerPart.Header.Descendants<Paragraph>().Count();
        }

        // 计算页脚中的段落
        foreach (var footerPart in mainPart.FooterParts)
        {
            count += footerPart.Footer.Descendants<Paragraph>().Count();
        }

        // 确保不为零
        return Math.Max(1, count);
    }

    /// <summary>
    /// 处理文档内容（正文、页眉、页脚等）
    /// </summary>
    private void ProcessDocumentContent(
        OpenXmlElement element,
        Dictionary<string, string> placeholderDict,
        ref int processedElements,
        int totalElements,
        IProgress<int> progress)
    {
        // 处理段落
        foreach (var paragraph in element.Descendants<Paragraph>())
        {
            ProcessParagraph(paragraph, placeholderDict);
            processedElements++;
            if (progress != null)
            {
                var progressValue = Math.Min(99, (int)((float)processedElements / totalElements * 100));
                progress.Report(progressValue);
            }
        }
    }

    /// <summary>
    /// 处理段落中的占位符替换
    /// </summary>
    private void ProcessParagraph(Paragraph paragraph, Dictionary<string, string> placeholderDict)
    {
        // 检查段落中是否有占位符 - 快速检查
        bool containsPlaceholder = false;
        string paragraphText = string.Join("", paragraph.Descendants<Text>().Select(t => t.Text));
        foreach (var placeholder in placeholderDict.Keys)
        {
            if (paragraphText.Contains(placeholder))
            {
                containsPlaceholder = true;
                break;
            }
        }

        if (!containsPlaceholder)
        {
            return; // 没有占位符，直接返回
        }

        // 处理简单情况：占位符完全在一个Run中
        bool allProcessed = ProcessSimplePlaceholders(paragraph, placeholderDict);

        // 如果还有未处理完的复杂情况（跨Run的占位符），使用复杂处理
        if (!allProcessed)
        {
            ProcessComplexPlaceholdersEnhanced(paragraph, placeholderDict);
        }
    }

    /// <summary>
    /// 处理简单情况：占位符完全在一个Run中
    /// </summary>
    /// <returns>是否所有占位符都已处理完毕</returns>
    private bool ProcessSimplePlaceholders(Paragraph paragraph, Dictionary<string, string> placeholderDict)
    {
        bool allProcessed = true;
        string paragraphText = string.Join("", paragraph.Descendants<Text>().Select(t => t.Text));

        // 查找未处理的占位符
        foreach (var placeholder in placeholderDict.Keys)
        {
            if (paragraphText.Contains(placeholder))
            {
                bool foundInRun = false;

                foreach (var run in paragraph.Descendants<Run>().ToList())
                {
                    string runText = string.Join("", run.Descendants<Text>().Select(t => t.Text));
                    if (runText.Contains(placeholder))
                    {
                        // 替换此Run中的占位符文本
                        foreach (var text in run.Descendants<Text>().ToList())
                        {
                            if (text.Text.Contains(placeholder))
                            {
                                text.Text = text.Text.Replace(placeholder, placeholderDict[placeholder]);
                                foundInRun = true;
                            }
                        }
                    }
                }

                // 如果占位符没有在任何一个Run中找到（可能跨Run），标记为未完成
                if (!foundInRun)
                {
                    allProcessed = false;
                }
            }
        }

        return allProcessed;
    }

    /// <summary>
    /// 处理复杂情况：占位符跨越多个Run - 增强版，解决长度不一致导致的格式错乱问题
    /// </summary>
    private void ProcessComplexPlaceholdersEnhanced(Paragraph paragraph, Dictionary<string, string> placeholderDict)
    {
        // 收集运行信息
        var allRuns = paragraph.Descendants<Run>().ToList();
        if (allRuns.Count == 0) return;

        // 提取段落全文
        string originalText = "";
        var runInfos = new List<RunInfo>();

        foreach (var run in allRuns)
        {
            string runText = string.Join("", run.Descendants<Text>().Select(t => t.Text));
            RunProperties runProps = run.RunProperties?.CloneNode(true) as RunProperties;

            runInfos.Add(new RunInfo
            {
                StartIndex = originalText.Length,
                EndIndex = originalText.Length + runText.Length,
                Text = runText,
                Properties = runProps
            });

            originalText += runText;
        }

        // 处理替换
        string resultText = originalText;
        var replacements = new List<TextReplacement>();

        foreach (var entry in placeholderDict)
        {
            string placeholder = entry.Key;
            string replacement = entry.Value;

            int startIndex = 0;
            while ((startIndex = resultText.IndexOf(placeholder, startIndex)) != -1)
            {
                replacements.Add(new TextReplacement
                {
                    StartIndex = startIndex,
                    EndIndex = startIndex + placeholder.Length,
                    OriginalText = placeholder,
                    NewText = replacement
                });

                startIndex += placeholder.Length;
            }
        }

        // 按照起始位置排序替换项（从后向前替换，避免索引变化）
        replacements = replacements.OrderByDescending(r => r.StartIndex).ToList();

        // 执行替换
        foreach (var replacement in replacements)
        {
            resultText = resultText.Remove(replacement.StartIndex, replacement.EndIndex - replacement.StartIndex)
                .Insert(replacement.StartIndex, replacement.NewText);
        }

        // 检查是否有变化
        if (resultText == originalText)
        {
            return; // 无变化，直接返回
        }

        // 重建段落内容
        var runsToRemove = paragraph.Elements<Run>().ToList();
        foreach (var run in runsToRemove)
        {
            run.Remove();
        }

        // 使用替换后的文本重新构建段落，保留每个区域的原始格式
        // 如果运行属性列表为空，则使用默认属性
        RunProperties defaultProps = allRuns.Count > 0 && allRuns[0].RunProperties != null
            ? allRuns[0].RunProperties.CloneNode(true) as RunProperties
            : new RunProperties();

        // 创建每个字符的格式映射
        var charFormatMap = new Dictionary<int, RunProperties>();

        foreach (var runInfo in runInfos)
        {
            RunProperties props = runInfo.Properties ?? defaultProps;
            for (int i = runInfo.StartIndex; i < runInfo.EndIndex; i++)
            {
                charFormatMap[i] = props;
            }
        }

        // 处理替换后文本的字符映射
        var adjustedFormatMap = new Dictionary<int, RunProperties>();
        int adjustment = 0;

        // 由于字符串长度变化，需要调整映射
        for (int i = 0; i < originalText.Length; i++)
        {
            // 查找该位置是否属于任何替换范围
            var replacementForPosition = replacements.FirstOrDefault(r => i >= r.StartIndex && i < r.EndIndex);

            if (replacementForPosition != null)
            {
                // 如果是替换区域的开始位置
                if (i == replacementForPosition.StartIndex)
                {
                    // 从替换开始位置获取格式
                    var formatProps = charFormatMap.ContainsKey(i) ? charFormatMap[i] : defaultProps;

                    // 对替换文本的每个字符应用该格式
                    for (int j = 0; j < replacementForPosition.NewText.Length; j++)
                    {
                        int newIndex = i + j + adjustment;
                        adjustedFormatMap[newIndex] = formatProps;
                    }

                    // 调整位置，跳过整个替换区域
                    adjustment += replacementForPosition.NewText.Length - replacementForPosition.OriginalText.Length;
                    i = replacementForPosition.EndIndex - 1; // -1因为for循环会++i
                }
            }
            else
            {
                // 不在替换区域内的字符，保留原格式
                int newIndex = i + adjustment;
                if (charFormatMap.ContainsKey(i))
                {
                    adjustedFormatMap[newIndex] = charFormatMap[i];
                }
            }
        }

        // 重组文本，根据格式分段
        int currentIndex = 0;
        while (currentIndex < resultText.Length)
        {
            // 获取当前字符的格式
            RunProperties currentProps = adjustedFormatMap.ContainsKey(currentIndex)
                ? adjustedFormatMap[currentIndex]
                : defaultProps;

            // 找出具有相同格式的连续字符
            int endIndex = currentIndex + 1;
            while (endIndex < resultText.Length &&
                   ArePropertiesEqual(
                       adjustedFormatMap.ContainsKey(endIndex) ? adjustedFormatMap[endIndex] : defaultProps,
                       currentProps))
            {
                endIndex++;
            }

            // 提取这段文本
            string segment = resultText.Substring(currentIndex, endIndex - currentIndex);

            // 创建新Run
            Run newRun = new Run();
            if (currentProps != null)
            {
                newRun.AppendChild(currentProps.CloneNode(true));
            }

            // 添加文本
            Text newText = new Text { Text = segment, Space = SpaceProcessingModeValues.Preserve };
            newRun.AppendChild(newText);
            paragraph.AppendChild(newRun);

            // 更新索引
            currentIndex = endIndex;
        }
    }

    /// <summary>
    /// 比较两个RunProperties是否相等
    /// </summary>
    private bool ArePropertiesEqual(RunProperties props1, RunProperties props2)
    {
        if (props1 == null && props2 == null) return true;
        if (props1 == null || props2 == null) return false;

        // 比较关键属性
        bool boldEqual = GetBoldValue(props1) == GetBoldValue(props2);
        bool italicEqual = GetItalicValue(props1) == GetItalicValue(props2);
        bool underlineEqual = GetUnderlineValue(props1) == GetUnderlineValue(props2);
        bool fontEqual = GetFontName(props1) == GetFontName(props2);
        bool sizeEqual = GetFontSize(props1) == GetFontSize(props2);
        bool colorEqual = GetColor(props1) == GetColor(props2);

        return boldEqual && italicEqual && underlineEqual && fontEqual && sizeEqual && colorEqual;
    }

    // 辅助方法：获取属性值
    private bool GetBoldValue(RunProperties props)
    {
        return props.Bold != null;
    }

    private bool GetItalicValue(RunProperties props)
    {
        return props.Italic != null;
    }

    private bool GetUnderlineValue(RunProperties props)
    {
        return props.Underline != null;
    }

    private string GetFontName(RunProperties props)
    {
        if (props.RunFonts != null && props.RunFonts.Ascii != null)
        {
            return props.RunFonts.Ascii.Value;
        }

        return "";
    }

    private string GetFontSize(RunProperties props)
    {
        if (props.FontSize != null && props.FontSize.Val != null)
        {
            return props.FontSize.Val.Value;
        }

        return "";
    }

    private string GetColor(RunProperties props)
    {
        if (props.Color != null && props.Color.Val != null)
        {
            return props.Color.Val.Value;
        }

        return "";
    }

    /// <summary>
    /// 运行信息类
    /// </summary>
    private class RunInfo
    {
        public int StartIndex { get; set; }
        public int EndIndex { get; set; }
        public string Text { get; set; }
        public RunProperties Properties { get; set; }
    }

    /// <summary>
    /// 文本替换信息类
    /// </summary>
    private class TextReplacement
    {
        public int StartIndex { get; set; }
        public int EndIndex { get; set; }
        public string OriginalText { get; set; }
        public string NewText { get; set; }
    }

    /// <summary>
    /// 从Word模板中提取所有占位符
    /// </summary>
    public async Task<List<string>> ExtractPlaceholdersAsync(string templatePath)
    {
        if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
        {
            return new List<string>();
        }

        return await Task.Run(() =>
        {
            var placeholders = new HashSet<string>();
            var regex = new Regex(@"\{([^{}]+)\}");

            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(templatePath, false))
                {
                    // 从主文档部分提取占位符
                    var mainPart = document.MainDocumentPart;
                    if (mainPart?.Document?.Body != null)
                    {
                        ExtractPlaceholdersFromPart(mainPart.Document.Body, placeholders, regex);
                    }

                    // 从页眉提取占位符
                    foreach (var headerPart in mainPart.HeaderParts)
                    {
                        ExtractPlaceholdersFromPart(headerPart.Header, placeholders, regex);
                    }

                    // 从页脚提取占位符
                    foreach (var footerPart in mainPart.FooterParts)
                    {
                        ExtractPlaceholdersFromPart(footerPart.Footer, placeholders, regex);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"提取Word占位符时出错: {ex.Message}");
            }

            return new List<string>(placeholders);
        });
    }

    /// <summary>
    /// 从OpenXML元素中提取占位符
    /// </summary>
    private void ExtractPlaceholdersFromPart(OpenXmlElement element, HashSet<string> placeholders, Regex regex)
    {
        // 收集所有文本
        var allTexts = element.Descendants<Text>();
        foreach (var text in allTexts)
        {
            var matches = regex.Matches(text.Text);
            foreach (Match match in matches)
            {
                placeholders.Add(match.Value); // 添加完整的占位符，包括大括号
            }
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