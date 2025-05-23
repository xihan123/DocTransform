using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocTransform.Constants;
using DocTransform.Models;
using DocTransform.Services;
using DocumentFormat.OpenXml.Packaging;
using Application = System.Windows.Application;
using Clipboard = System.Windows.Clipboard;
using DataFormats = System.Windows.DataFormats;
using DragEventArgs = System.Windows.DragEventArgs;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Timer = System.Threading.Timer;

namespace DocTransform.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ExcelService _excelService;

    // Excel模板相关属性
    private readonly ExcelTemplateService _excelTemplateService;
    private readonly IdCardService _idCardService;
    private readonly ImageProcessingService _imageProcessingService;
    private readonly WordService _wordService;

    [ObservableProperty] private ObservableCollection<string> _availableColumns = new();

    [ObservableProperty] private ObservableCollection<string> _availableIdCardColumns = new();

    [ObservableProperty] private ObservableCollection<string> _availableKeyColumns = new();

    [ObservableProperty] private ExcelData _currentExcelData;

    [ObservableProperty] private ObservableCollection<string> _detectedExcelPlaceholders = new();

    // 身份证信息提取相关属性
    [ObservableProperty] private bool _enableIdCardExtraction;

    private ExcelData _excelData = new();

    [ObservableProperty] private string _excelFilePath = string.Empty;

    [ObservableProperty] private string _excelTemplatePath = string.Empty;

    // 身份证占位符集合
    [ObservableProperty] private List<string> _idCardPlaceholders = PlaceholderConstants.AllPlaceholders;

    /// <summary>
    ///     图片目录列表
    /// </summary>
    [ObservableProperty] private ObservableCollection<ImageSourceDirectory> _imageDirectories = new();

    /// <summary>
    ///     图片填充模式
    /// </summary>
    [ObservableProperty] private ImageFillMode _imageFillMode = ImageFillMode.Fit;

    /// <summary>
    ///     可选填充模式列表
    /// </summary>
    [ObservableProperty] private List<ImageFillModeItem> _imageFillModeItems = ImageFillModeItem.GetAll();

    /// <summary>
    ///     图片填满单元格的程度（百分比）
    /// </summary>
    [ObservableProperty] private int _imageFillPercentage = 90;

    [ObservableProperty] private bool _isMultiTableMode;

    [ObservableProperty] private bool _isProcessing;

    // 多表格相关属性
    [ObservableProperty] private MultiTableData _multiTableData = new();

    [ObservableProperty] private string _outputDirectory = string.Empty;

    [ObservableProperty] private string _outputFileNameTemplate = "{序号}_{姓名}_{时间}";

    [ObservableProperty] private int _processedItems;

    [ObservableProperty] private string _processResultText = string.Empty;

    [ObservableProperty] private bool _processSuccess;

    [ObservableProperty] private int _progressValue;

    [ObservableProperty] private ObservableCollection<string> _selectedColumns = new();

    [ObservableProperty] private string _selectedIdCardColumn = string.Empty;

    /// <summary>
    ///     当前选择的填充模式项
    /// </summary>
    [ObservableProperty] private ImageFillModeItem _selectedImageFillModeItem;

    /// <summary>
    ///     当前选中的用于匹配图片的列名
    /// </summary>
    [ObservableProperty] private string _selectedImageMatchingColumn = string.Empty;

    [ObservableProperty] private string _selectedKeyColumn = string.Empty;

    [ObservableProperty] private string _statusMessage = "准备就绪";

    [ObservableProperty] private int _totalItems;

    [ObservableProperty] private bool _useExcelTemplate;

    /// <summary>
    ///     是否启用图片替换功能
    /// </summary>
    [ObservableProperty] private bool _useImageReplacement;

    [ObservableProperty] private string _wordTemplatePath = string.Empty;

    public MainViewModel()
    {
        _excelService = new ExcelService();
        _wordService = new WordService();
        _idCardService = new IdCardService();
        _imageProcessingService = new ImageProcessingService();
        _excelTemplateService = new ExcelTemplateService(_imageProcessingService);

        // 设置默认输出目录为"我的文档"
        OutputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        // 初始化选中的填充模式
        SelectedImageFillModeItem = ImageFillModeItems.First(item => item.Value == ImageFillMode.Fit);
    }

    [RelayCommand]
    private async Task BrowseExcelFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel文件 (*.xlsx)|*.xlsx",
            Title = "选择Excel文件"
        };

        if (dialog.ShowDialog() == true)
        {
            ExcelFilePath = dialog.FileName;

            if (IsMultiTableMode)
                await AddExcelFileToMultiTable(ExcelFilePath);
            else
                await LoadSingleExcelFile(ExcelFilePath);
        }
    }

    // 单文件模式加载Excel
    private async Task LoadSingleExcelFile(string filePath)
    {
        try
        {
            StatusMessage = "正在加载Excel文件...";
            IsProcessing = true;

            _excelData = await _excelService.ReadExcelFileAsync(filePath);
            CurrentExcelData = _excelData;

            // 更新可用列
            UpdateAvailableColumns();

            // 更新身份证列
            UpdateIdCardColumns();

            StatusMessage = $"Excel文件加载完成，共有 {_excelData.Rows.Count} 行数据";
        }
        catch (Exception ex)
        {
            StatusMessage = $"加载Excel文件失败: {ex.Message}";
            ExcelFilePath = string.Empty;
        }
        finally
        {
            IsProcessing = false;
        }
    }

    // 多表格模式添加Excel文件
    private async Task AddExcelFileToMultiTable(string filePath)
    {
        try
        {
            StatusMessage = "正在加载Excel文件...";
            IsProcessing = true;

            var allSheets = await _excelService.ReadAllSheetsAsync(filePath);

            if (allSheets.Count == 0)
            {
                StatusMessage = "Excel文件中没有有效数据";
                return;
            }

            // 添加到多表格数据中
            foreach (var sheet in allSheets) _multiTableData.Tables.Add(sheet);

            // 更新可用的键列（必须存在于所有表中）
            UpdateKeyColumns();

            // 更新可用列（所有表中的列，无重复）
            UpdateAvailableColumns();

            // 更新身份证列
            UpdateIdCardColumns();

            StatusMessage = $"添加了 {allSheets.Count} 个工作表，共有 {_multiTableData.TotalRowCount} 行数据";
        }
        catch (Exception ex)
        {
            StatusMessage = $"加载Excel文件失败: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
        }
    }

    // 更新可用列
    private void UpdateAvailableColumns()
    {
        AvailableColumns.Clear();
        SelectedColumns.Clear();

        if (IsMultiTableMode)
        {
            foreach (var header in _multiTableData.AllHeaders) AvailableColumns.Add(header);

            // 如果选择了Key列，则执行合并
            if (!string.IsNullOrEmpty(SelectedKeyColumn)) _multiTableData.MergeData(SelectedKeyColumn);
        }
        else if (CurrentExcelData != null)
        {
            foreach (var header in CurrentExcelData.Headers) AvailableColumns.Add(header);
        }
    }

    // 更新可用于匹配的键列
    private void UpdateKeyColumns()
    {
        AvailableKeyColumns.Clear();
        SelectedKeyColumn = string.Empty;

        if (IsMultiTableMode)
            // 只有在所有表中都存在的列才能作为键列
            foreach (var header in _multiTableData.CommonHeaders)
            {
                AvailableKeyColumns.Add(header);

                // 自动选择可能的键列
                if (string.IsNullOrEmpty(SelectedKeyColumn))
                    if (header.Contains("身份证", StringComparison.OrdinalIgnoreCase) ||
                        header.Contains("ID", StringComparison.OrdinalIgnoreCase) ||
                        header.Contains("编号", StringComparison.OrdinalIgnoreCase) ||
                        header.Contains("姓名", StringComparison.OrdinalIgnoreCase) ||
                        header.Contains("名字", StringComparison.OrdinalIgnoreCase))
                        SelectedKeyColumn = header;
            }
    }

    // 更新可能的身份证列
    private void UpdateIdCardColumns()
    {
        AvailableIdCardColumns.Clear();
        SelectedIdCardColumn = string.Empty;

        List<string> headers;

        if (IsMultiTableMode)
            // 多表格模式下，使用合并后的所有列
            headers = _multiTableData.AllHeaders;
        else if (CurrentExcelData != null)
            // 单表格模式下，使用当前表格的列
            headers = CurrentExcelData.Headers;
        else
            return;

        foreach (var header in headers)
            // 检测可能的身份证列（名称中包含"身份证"、"证件"等关键词）
            if (header.Contains("身份证", StringComparison.OrdinalIgnoreCase) ||
                header.Contains("证件", StringComparison.OrdinalIgnoreCase) ||
                header.Contains("ID", StringComparison.OrdinalIgnoreCase))
            {
                AvailableIdCardColumns.Add(header);
                if (string.IsNullOrEmpty(SelectedIdCardColumn))
                    // 优先选择包含"身份证"的列作为默认选择
                    if (header.Contains("身份证", StringComparison.OrdinalIgnoreCase))
                        SelectedIdCardColumn = header;
            }

        // 如果没有找到包含"身份证"的列，则使用第一个可用的身份证列
        if (string.IsNullOrEmpty(SelectedIdCardColumn) && AvailableIdCardColumns.Count > 0)
            SelectedIdCardColumn = AvailableIdCardColumns[0];

        // 如果启用了身份证信息提取但没有找到可用的身份证列，可以考虑给用户一个提示
        if (EnableIdCardExtraction && string.IsNullOrEmpty(SelectedIdCardColumn))
            StatusMessage = "警告：已启用身份证信息提取，但未找到可能的身份证列";
    }

    // 切换单表格/多表格模式
    [RelayCommand]
    private void ToggleTableMode()
    {
        // 添加调试输出，确认命令被触发
        Debug.WriteLine($"ToggleTableMode 被调用，当前模式: {IsMultiTableMode}");

        IsMultiTableMode = !IsMultiTableMode;

        if (IsMultiTableMode)
        {
            // 切换到多表格模式
            _multiTableData.Clear();
            if (CurrentExcelData != null)
            {
                // 将当前的单表格添加到多表格中
                _multiTableData.Tables.Add(CurrentExcelData);
                UpdateKeyColumns();
            }

            StatusMessage = "已切换到多表格模式";
        }
        else
        {
            // 切换到单表格模式
            CurrentExcelData = _multiTableData.Tables.Count > 0 ? _multiTableData.Tables[0] : new ExcelData();
            _excelData = CurrentExcelData;
            UpdateAvailableColumns();
            UpdateIdCardColumns();
            StatusMessage = "已切换到单表格模式";
        }

        // 添加调试输出，确认模式已切换
        Debug.WriteLine($"模式已切换为: {(IsMultiTableMode ? "多表格" : "单表格")}");
    }

    // 移除指定表格
    [RelayCommand]
    private void RemoveTable(ExcelData table)
    {
        if (_multiTableData.Tables.Contains(table))
        {
            _multiTableData.Tables.Remove(table);
            UpdateKeyColumns();
            UpdateAvailableColumns();
            UpdateIdCardColumns();
            StatusMessage = $"已移除表格: {table.SourceFileName}";
        }
    }

    // 修改键列选择事件
    partial void OnSelectedKeyColumnChanged(string value)
    {
        if (IsMultiTableMode && !string.IsNullOrEmpty(value))
        {
            _multiTableData.MergeData(value);
            StatusMessage = $"已使用 \"{value}\" 列合并数据，共有 {_multiTableData.MergedRows.Count} 条记录";
        }
    }

    [RelayCommand]
    private async Task BrowseWordTemplate()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Word文档 (*.docx)|*.docx",
            Title = "选择Word模板"
        };

        if (dialog.ShowDialog() == true)
        {
            WordTemplatePath = dialog.FileName;

            // 验证Word模板有效性
            var isValid = await _wordService.IsValidTemplateAsync(WordTemplatePath);
            if (isValid)
            {
                // 自动检查占位符
                await CheckPlaceholders();
            }
            else
            {
                StatusMessage = "选择的Word模板无效";
                WordTemplatePath = string.Empty;
            }
        }
    }

    [RelayCommand]
    private void BrowseOutputDirectory()
    {
        var dialog = new FolderBrowserDialog
        {
            Description = "选择输出目录",
            UseDescriptionForTitle = true
        };

        if (dialog.ShowDialog() == DialogResult.OK) OutputDirectory = dialog.SelectedPath;
    }

    [RelayCommand]
    public async Task HandleFileDrop(DragEventArgs e)
    {
        try
        {
            // 确保使用正确的数据格式
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (files != null && files.Length > 0)
                {
                    var file = files[0];
                    Debug.WriteLine($"拖放文件: {file}");

                    if (Path.GetExtension(file).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        ExcelFilePath = file;
                        await LoadExcelFile(file);
                    }
                    else if (Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        WordTemplatePath = file;

                        // 验证Word模板有效性
                        var isValid = await _wordService.IsValidTemplateAsync(WordTemplatePath);
                        if (!isValid)
                        {
                            StatusMessage = "选择的Word模板无效";
                            WordTemplatePath = string.Empty;
                        }
                        else
                        {
                            // 自动检查占位符
                            await CheckPlaceholders();
                        }
                    }
                    else
                    {
                        StatusMessage = "不支持的文件类型";
                    }
                }
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"处理拖放文件时出错: {ex.Message}";
            Debug.WriteLine($"拖放异常: {ex}");
        }
    }

    [RelayCommand]
    private void MoveColumnToSelected(string column)
    {
        if (AvailableColumns.Contains(column))
        {
            AvailableColumns.Remove(column);
            SelectedColumns.Add(column);
            _excelData.SelectedColumns.Add(column);
        }
    }

    [RelayCommand]
    private void MoveColumnToAvailable(string column)
    {
        if (SelectedColumns.Contains(column))
        {
            SelectedColumns.Remove(column);
            AvailableColumns.Add(column);
            _excelData.SelectedColumns.Remove(column);
        }
    }

    [RelayCommand]
    private void MoveAllColumnsToSelected()
    {
        var columnsToMove = AvailableColumns.ToList();
        foreach (var column in columnsToMove)
        {
            SelectedColumns.Add(column);
            _excelData.SelectedColumns.Add(column);
        }

        AvailableColumns.Clear();
    }

    [RelayCommand]
    private void MoveAllColumnsToAvailable()
    {
        var columnsToMove = SelectedColumns.ToList();
        foreach (var column in columnsToMove) AvailableColumns.Add(column);
        SelectedColumns.Clear();
        _excelData.SelectedColumns.Clear();
    }

    [RelayCommand]
    private void CopyToClipboard(string text)
    {
        try
        {
            Clipboard.SetText(text);
            // 显示一个临时提示，表示文本已复制
            StatusMessage = $"已复制「{text}」到剪贴板";

            // 使用定时器在几秒后恢复原始状态消息
            var timer = new Timer(_ =>
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (StatusMessage.StartsWith($"已复制「{text}」")) StatusMessage = "准备就绪";
                });
            }, null, 3000, Timeout.Infinite);
        }
        catch (Exception ex)
        {
            StatusMessage = $"复制到剪贴板失败: {ex.Message}";
        }
    }

    /// <summary>
    ///     生成文档
    /// </summary>
    [RelayCommand]
    private async Task GenerateDocuments()
    {
        // 验证基本条件...
        if (string.IsNullOrEmpty(OutputDirectory) || !Directory.Exists(OutputDirectory))
        {
            StatusMessage = "请选择有效的输出目录";
            return;
        }

        if (string.IsNullOrEmpty(OutputFileNameTemplate)) OutputFileNameTemplate = "{序号}_{时间}";

        // 验证模板和功能选择
        var hasWordTemplate = !string.IsNullOrEmpty(WordTemplatePath) && File.Exists(WordTemplatePath);
        var hasExcelTemplate = UseExcelTemplate && !string.IsNullOrEmpty(ExcelTemplatePath) &&
                               File.Exists(ExcelTemplatePath);
        var hasImageDirectories = UseImageReplacement && ImageDirectories.Count > 0;

        if (!hasWordTemplate && !hasExcelTemplate)
        {
            StatusMessage = "请至少选择一个Word模板或Excel模板";
            return;
        }

        if (hasImageDirectories && !hasExcelTemplate)
        {
            StatusMessage = "图片替换功能需要启用Excel模板";
            return;
        }

        // 验证数据源...
        List<Dictionary<string, string>> dataRows;

        if (IsMultiTableMode)
        {
            // 多表格模式处理...
            if (_multiTableData.Tables.Count == 0)
            {
                StatusMessage = "请先添加Excel文件";
                return;
            }

            if (string.IsNullOrEmpty(SelectedKeyColumn))
            {
                StatusMessage = "请选择用于匹配记录的列";
                return;
            }

            dataRows = _multiTableData.MergedRows;

            if (dataRows.Count == 0)
            {
                StatusMessage = "合并后没有有效数据";
                return;
            }
        }
        else
        {
            if (_excelData.Rows.Count == 0)
            {
                StatusMessage = "Excel文件中没有有效数据";
                return;
            }

            dataRows = _excelData.Rows;
        }

        // 检查图片目录的匹配列设置
        if (hasImageDirectories)
        {
            var directoriesWithoutColumn = ImageDirectories
                .Where(d => string.IsNullOrEmpty(d.MatchingColumn))
                .ToList();

            if (directoriesWithoutColumn.Any())
            {
                StatusMessage =
                    $"图片目录 {string.Join(", ", directoriesWithoutColumn.Select(d => d.DirectoryName))} 未设置匹配列";
                return;
            }
        }

        // 其他校验...
        if (EnableIdCardExtraction && string.IsNullOrEmpty(SelectedIdCardColumn))
        {
            StatusMessage = "已启用身份证信息提取，但未选择身份证列";
            return;
        }

        try
        {
            IsProcessing = true;
            ProgressValue = 0;
            TotalItems = dataRows.Count;
            ProcessedItems = 0;
            ProcessResultText = string.Empty;

            var successCount = 0;
            var failCount = 0;

            // 计算总处理任务数
            var totalTasks = dataRows.Count * (hasWordTemplate ? 1 : 0) +
                             dataRows.Count * (hasExcelTemplate ? 1 : 0);
            var completedTasks = 0;

            for (var i = 0; i < dataRows.Count; i++)
            {
                var rowData = new Dictionary<string, string>(dataRows[i]);

                // 添加一些特殊变量
                rowData["序号"] = (i + 1).ToString();
                rowData["时间"] = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                rowData["日期"] = DateTime.Now.ToString("yyyy-MM-dd");

                // 处理身份证信息提取
                if (EnableIdCardExtraction && !string.IsNullOrEmpty(SelectedIdCardColumn) &&
                    rowData.TryGetValue(SelectedIdCardColumn, out var idCard) &&
                    !string.IsNullOrEmpty(idCard))
                    try
                    {
                        // 提取身份证信息...
                        var gender = _idCardService.ExtractGender(idCard);
                        var birthDate = _idCardService.ExtractBirthDate(idCard);
                        var region = _idCardService.ExtractRegion(idCard);

                        rowData["身份证性别"] = gender;
                        rowData["身份证出生日期"] = birthDate;
                        rowData["身份证籍贯"] = region;

                        rowData["身份证出生年"] = _idCardService.ExtractBirthDate(idCard, "yyyy");
                        rowData["身份证出生月"] = _idCardService.ExtractBirthDate(idCard, "MM");
                        rowData["身份证出生日"] = _idCardService.ExtractBirthDate(idCard, "dd");
                        rowData["身份证年龄"] = CalculateAge(birthDate);
                    }
                    catch (Exception ex)
                    {
                        // 处理异常...
                        Debug.WriteLine($"身份证信息提取出错: {ex.Message}");

                        rowData["身份证性别"] = "未知";
                        rowData["身份证出生日期"] = "未知";
                        rowData["身份证籍贯"] = "未知";
                        rowData["身份证出生年"] = "未知";
                        rowData["身份证出生月"] = "未知";
                        rowData["身份证出生日"] = "未知";
                        rowData["身份证年龄"] = "未知";
                    }

                // 生成文件名
                var fileName = OutputFileNameTemplate;
                foreach (var item in rowData)
                    fileName = fileName.Replace($"{{{item.Key}}}", item.Value ?? string.Empty,
                        StringComparison.OrdinalIgnoreCase);

                // 处理无效字符
                foreach (var invalidChar in Path.GetInvalidFileNameChars())
                    fileName = fileName.Replace(invalidChar, '_');

                // 确保文件名有效
                if (string.IsNullOrWhiteSpace(fileName) || fileName.All(c => c == '_'))
                    fileName = $"Document_{i + 1}_{DateTime.Now:yyyyMMdd-HHmmss}";

                // 处理Word模板
                if (hasWordTemplate)
                {
                    var wordOutputPath = Path.Combine(OutputDirectory, $"{fileName}.docx");

                    // 异步处理Word文档
                    var wordProgress = new Progress<int>(value =>
                    {
                        ProgressValue = (completedTasks * 100 + value) / totalTasks;
                    });

                    var wordResult =
                        await _wordService.ProcessTemplateAsync(WordTemplatePath, wordOutputPath, rowData,
                            wordProgress);

                    if (wordResult.Success)
                    {
                        successCount++;
                    }
                    else
                    {
                        failCount++;
                        StatusMessage = $"处理第 {i + 1} 行Word文档时出错: {wordResult.Message}";
                    }

                    completedTasks++;
                }

                // 处理Excel模板
                if (hasExcelTemplate)
                {
                    var excelOutputPath = Path.Combine(OutputDirectory, $"{fileName}.xlsx");

                    // 异步处理Excel文档
                    var excelProgress = new Progress<int>(value =>
                    {
                        ProgressValue = (completedTasks * 100 + value) / totalTasks;
                    });

                    (bool Success, string Message) excelResult;

                    if (hasImageDirectories)
                        // 使用带图片处理的方法
                        excelResult = await _excelTemplateService.ProcessTemplateWithImagesAsync(
                            ExcelTemplatePath,
                            excelOutputPath,
                            rowData,
                            ImageDirectories,
                            ImageFillMode, // 直接使用枚举值
                            ImageFillPercentage,
                            excelProgress);
                    else
                        // 使用普通的模板处理方法
                        excelResult = await _excelTemplateService.ProcessTemplateAsync(
                            ExcelTemplatePath,
                            excelOutputPath,
                            rowData,
                            excelProgress);

                    if (excelResult.Success)
                    {
                        successCount++;
                    }
                    else
                    {
                        failCount++;
                        StatusMessage = $"处理第 {i + 1} 行Excel文档时出错: {excelResult.Message}";
                    }

                    completedTasks++;
                }

                ProcessedItems = i + 1;
            }

            // 处理完成
            ProcessResultText = $"处理完成：成功 {successCount} 个，失败 {failCount} 个";
            ProcessSuccess = failCount == 0;
            StatusMessage = $"文档生成完成，输出到 {OutputDirectory}";
            ProgressValue = 100;
        }
        catch (Exception ex)
        {
            StatusMessage = $"处理过程中出错: {ex.Message}";
            ProcessSuccess = false;
            ProcessResultText = $"处理失败: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
        }
    }

    /// <summary>
    ///     检查并记录Word模板中的占位符
    /// </summary>
    [RelayCommand]
    private async Task CheckPlaceholders()
    {
        if (string.IsNullOrEmpty(WordTemplatePath) || !File.Exists(WordTemplatePath))
        {
            StatusMessage = "请先选择有效的Word模板";
            return;
        }

        try
        {
            IsProcessing = true;
            StatusMessage = "正在检查Word模板中的占位符...";

            await Task.Run(() =>
            {
                try
                {
                    // 创建临时副本
                    var tempFile = Path.GetTempFileName() + ".docx";
                    File.Copy(WordTemplatePath, tempFile, true);

                    using (var doc = WordprocessingDocument.Open(tempFile, false))
                    {
                        if (doc.MainDocumentPart?.Document?.Body != null)
                        {
                            var docText = doc.MainDocumentPart.Document.Body.InnerText;
                            var placeholders = new HashSet<string>();

                            // 使用正则表达式匹配所有 {xxx} 格式的占位符
                            var matches = Regex.Matches(docText, @"\{([^{}]+)\}");

                            foreach (Match match in matches) placeholders.Add(match.Value);

                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (placeholders.Count > 0)
                                {
                                    StatusMessage = $"模板中找到 {placeholders.Count} 个占位符";

                                    // 检查身份证相关占位符
                                    var idCardPlaceholders = PlaceholderConstants.AllPlaceholders;
                                    var foundIdCardPlaceholders = false;

                                    foreach (var idCardPlaceholder in idCardPlaceholders)
                                        if (placeholders.Contains(idCardPlaceholder))
                                        {
                                            foundIdCardPlaceholders = true;
                                            break;
                                        }

                                    if (foundIdCardPlaceholders && !EnableIdCardExtraction)
                                    {
                                        // 自动启用身份证信息提取功能
                                        EnableIdCardExtraction = true;
                                        StatusMessage += " (已自动启用身份证信息提取功能)";
                                    }
                                }
                                else
                                {
                                    StatusMessage = "模板中未找到任何占位符";
                                }
                            });
                        }
                    }

                    // 清理临时文件
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                    }
                }
                catch (Exception ex)
                {
                    Application.Current.Dispatcher.Invoke(() => { StatusMessage = $"检查占位符时出错: {ex.Message}"; });
                }
            });
        }
        catch (Exception ex)
        {
            StatusMessage = $"检查占位符失败: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
        }
    }

    private async Task LoadExcelFile(string filePath)
    {
        try
        {
            StatusMessage = "正在加载Excel文件...";
            IsProcessing = true;

            _excelData = await _excelService.ReadExcelFileAsync(filePath);

            // 更新可用列
            AvailableColumns.Clear();
            SelectedColumns.Clear();
            _excelData.SelectedColumns.Clear();

            // 更新可能的身份证列
            AvailableIdCardColumns.Clear();
            SelectedIdCardColumn = string.Empty;

            foreach (var header in _excelData.Headers)
            {
                AvailableColumns.Add(header);

                // 检测可能的身份证列（名称中包含"身份证"、"证件"等关键词）
                if (header.Contains("身份证", StringComparison.OrdinalIgnoreCase) ||
                    header.Contains("证件", StringComparison.OrdinalIgnoreCase) ||
                    header.Contains("ID", StringComparison.OrdinalIgnoreCase))
                {
                    AvailableIdCardColumns.Add(header);
                    if (string.IsNullOrEmpty(SelectedIdCardColumn)) SelectedIdCardColumn = header;
                }
            }

            StatusMessage = $"Excel文件加载完成，共有 {_excelData.Rows.Count} 行数据";
        }
        catch (Exception ex)
        {
            StatusMessage = $"加载Excel文件失败: {ex.Message}";
            ExcelFilePath = string.Empty;
        }
        finally
        {
            IsProcessing = false;
        }
    }

    [RelayCommand]
    private void OpenOutputFolder()
    {
        if (Directory.Exists(OutputDirectory))
            try
            {
                Process.Start("explorer.exe", OutputDirectory);
            }
            catch (Exception ex)
            {
                StatusMessage = $"打开输出文件夹失败: {ex.Message}";
            }
        else
            StatusMessage = "输出目录不存在";
    }

    // 计算年龄的辅助方法
    private string CalculateAge(string birthDateStr)
    {
        try
        {
            if (DateTime.TryParse(birthDateStr, out var birthDate))
            {
                var age = DateTime.Today.Year - birthDate.Year;
                if (birthDate.Date > DateTime.Today.AddYears(-age)) age--;
                return age.ToString();
            }
        }
        catch
        {
        }

        return "未知";
    }

    /// <summary>
    ///     浏览Excel模板文件
    /// </summary>
    [RelayCommand]
    private async Task BrowseExcelTemplate()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel模板 (*.xlsx)|*.xlsx",
            Title = "选择Excel模板"
        };

        if (dialog.ShowDialog() == true)
        {
            ExcelTemplatePath = dialog.FileName;

            // 验证Excel模板有效性
            var isValid = await _excelTemplateService.IsValidTemplateAsync(ExcelTemplatePath);
            if (!isValid)
            {
                StatusMessage = "选择的Excel模板无效";
                ExcelTemplatePath = string.Empty;
            }
            else
            {
                // 自动检查Excel占位符
                await CheckExcelPlaceholders();
            }
        }
    }

    /// <summary>
    ///     检查Excel模板中的占位符
    /// </summary>
    [RelayCommand]
    private async Task CheckExcelPlaceholders()
    {
        if (string.IsNullOrEmpty(ExcelTemplatePath) || !File.Exists(ExcelTemplatePath))
        {
            StatusMessage = "请先选择有效的Excel模板";
            return;
        }

        try
        {
            IsProcessing = true;
            StatusMessage = "正在检查Excel模板中的占位符...";

            var placeholders = await _excelTemplateService.ExtractPlaceholdersAsync(ExcelTemplatePath);

            DetectedExcelPlaceholders.Clear();
            foreach (var placeholder in placeholders) DetectedExcelPlaceholders.Add(placeholder);

            if (placeholders.Count > 0)
                StatusMessage = $"Excel模板中找到 {placeholders.Count} 个占位符";
            else
                StatusMessage = "Excel模板中未找到任何占位符";
        }
        catch (Exception ex)
        {
            StatusMessage = $"检查Excel占位符失败: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
        }
    }

    /// <summary>
    ///     清除Excel模板
    /// </summary>
    [RelayCommand]
    private void ClearExcelTemplate()
    {
        ExcelTemplatePath = string.Empty;
        DetectedExcelPlaceholders.Clear();
        UseExcelTemplate = false;
    }

    /// <summary>
    ///     添加图片目录
    /// </summary>
    [RelayCommand]
    private async Task AddImageDirectory()
    {
        var dialog = new FolderBrowserDialog
        {
            Description = "选择包含图片的目录",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
            try
            {
                IsProcessing = true;

                var directoryPath = dialog.SelectedPath;
                var directoryName = Path.GetFileName(directoryPath);

                // 确保目录名称不为空
                if (string.IsNullOrEmpty(directoryName)) directoryName = new DirectoryInfo(directoryPath).Name;

                // 检查是否已存在同名目录
                if (ImageDirectories.Any(d =>
                        d.DirectoryName.Equals(directoryName, StringComparison.OrdinalIgnoreCase)))
                {
                    var suffix = 1;
                    var originalName = directoryName;
                    // 自动添加数字后缀
                    while (ImageDirectories.Any(d =>
                               d.DirectoryName.Equals(directoryName, StringComparison.OrdinalIgnoreCase)))
                        directoryName = $"{originalName}_{suffix++}";
                }

                StatusMessage = "正在扫描图片目录...";

                // 扫描目录中的图片
                var imageFiles = await _imageProcessingService.ScanDirectoryForImagesAsync(directoryPath);

                if (imageFiles.Count == 0)
                {
                    StatusMessage = $"目录 {directoryName} 中未找到支持的图片文件";
                    return;
                }

                // 创建图片目录对象
                var imageDirectory = new ImageSourceDirectory
                {
                    DirectoryPath = directoryPath,
                    DirectoryName = directoryName,
                    MatchingColumn = SelectedImageMatchingColumn,
                    ImageFiles = imageFiles
                };

                // 添加到列表
                ImageDirectories.Add(imageDirectory);

                // 自动启用图片替换功能
                UseImageReplacement = true;

                StatusMessage = $"已添加图片目录: {directoryName}，包含 {imageFiles.Count} 个图片";
            }
            catch (Exception ex)
            {
                StatusMessage = $"添加图片目录失败: {ex.Message}";
            }
            finally
            {
                IsProcessing = false;
            }
    }

    /// <summary>
    ///     移除图片目录
    /// </summary>
    /// <param name="directory">要移除的目录</param>
    [RelayCommand]
    private void RemoveImageDirectory(ImageSourceDirectory directory)
    {
        if (ImageDirectories.Contains(directory))
        {
            ImageDirectories.Remove(directory);
            StatusMessage = $"已移除图片目录: {directory.DirectoryName}";

            // 如果没有图片目录，自动禁用图片替换功能
            if (ImageDirectories.Count == 0) UseImageReplacement = false;
        }
    }

    /// <summary>
    ///     设置图片目录的匹配列
    /// </summary>
    /// <param name="parameters">参数元组 (ImageSourceDirectory, string)</param>
    [RelayCommand]
    private void SetDirectoryMatchingColumn(object parameters)
    {
        if (parameters is ValueTuple<ImageSourceDirectory, string> tuple)
        {
            var (directory, columnName) = tuple;

            if (directory != null && !string.IsNullOrEmpty(columnName))
            {
                directory.MatchingColumn = columnName;

                // 触发UI更新
                var index = ImageDirectories.IndexOf(directory);
                if (index >= 0) ImageDirectories[index] = directory;

                StatusMessage = $"已设置 {directory.DirectoryName} 的匹配列为 {columnName}";
            }
        }
    }

    /// <summary>
    ///     检查指定目录中的图片文件
    /// </summary>
    /// <param name="directory">图片目录对象</param>
    [RelayCommand]
    private async Task CheckDirectoryImages(ImageSourceDirectory directory)
    {
        if (directory == null) return;

        try
        {
            IsProcessing = true;
            StatusMessage = $"正在扫描目录 {directory.DirectoryName} 中的图片...";

            // 重新扫描目录
            var imageFiles = await _imageProcessingService.ScanDirectoryForImagesAsync(directory.DirectoryPath);

            // 更新图片列表
            directory.ImageFiles.Clear();
            foreach (var file in imageFiles) directory.ImageFiles.Add(file);

            // 触发UI更新
            var index = ImageDirectories.IndexOf(directory);
            if (index >= 0) ImageDirectories[index] = directory;

            StatusMessage = $"目录 {directory.DirectoryName} 中找到 {imageFiles.Count} 个图片";
        }
        catch (Exception ex)
        {
            StatusMessage = $"扫描图片目录失败: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
        }
    }

    /// <summary>
    ///     清除所有图片目录
    /// </summary>
    [RelayCommand]
    private void ClearAllImageDirectories()
    {
        ImageDirectories.Clear();
        UseImageReplacement = false;
        StatusMessage = "已清除所有图片目录";
    }

    // 更新数据列变更处理
    partial void OnAvailableColumnsChanged(ObservableCollection<string> value)
    {
        // 当可用列更新时，也更新图片匹配列下拉框
        // 保持当前选择（如果存在）
        if (!string.IsNullOrEmpty(SelectedImageMatchingColumn) &&
            value.Contains(SelectedImageMatchingColumn))
        {
            // 保持当前选择
        }
        else if (value.Count > 0)
        {
            // 尝试智能选择匹配列（优先选择名称、ID等常用标识符）
            var preferredColumns = value.Where(c =>
                c.Contains("姓名", StringComparison.OrdinalIgnoreCase) ||
                c.Contains("名字", StringComparison.OrdinalIgnoreCase) ||
                c.Contains("ID", StringComparison.OrdinalIgnoreCase) ||
                c.Contains("编号", StringComparison.OrdinalIgnoreCase) ||
                c.Contains("身份证", StringComparison.OrdinalIgnoreCase)
            ).ToList();

            if (preferredColumns.Any())
                SelectedImageMatchingColumn = preferredColumns.First();
            else
                // 默认选择第一列
                SelectedImageMatchingColumn = value.First();
        }
        else
        {
            SelectedImageMatchingColumn = string.Empty;
        }
    }

    // 当选中的填充模式项变化时，更新实际的枚举值
    partial void OnSelectedImageFillModeItemChanged(ImageFillModeItem value)
    {
        if (value != null) ImageFillMode = value.Value;
    }
}