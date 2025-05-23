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
    private readonly IdCardService _idCardService;
    private readonly WordService _wordService;

    [ObservableProperty] private ObservableCollection<string> _availableColumns = new();

    [ObservableProperty] private ObservableCollection<string> _availableIdCardColumns = new();

    [ObservableProperty] private ObservableCollection<string> _availableKeyColumns = new();

    [ObservableProperty] private ExcelData _currentExcelData;

    // 身份证信息提取相关属性
    [ObservableProperty] private bool _enableIdCardExtraction;

    private ExcelData _excelData = new();

    [ObservableProperty] private string _excelFilePath = string.Empty;

    // 身份证占位符集合
    [ObservableProperty] private List<string> _idCardPlaceholders = PlaceholderConstants.AllPlaceholders;

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

    [ObservableProperty] private string _selectedKeyColumn = string.Empty;

    [ObservableProperty] private string _statusMessage = "准备就绪";

    [ObservableProperty] private int _totalItems;

    [ObservableProperty] private string _wordTemplatePath = string.Empty;

    public MainViewModel()
    {
        _excelService = new ExcelService();
        _wordService = new WordService();
        _idCardService = new IdCardService();

        // 设置默认输出目录为"我的文档"
        OutputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
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

    [RelayCommand]
    private async Task GenerateDocuments()
    {
        if (string.IsNullOrEmpty(WordTemplatePath))
        {
            StatusMessage = "请选择Word模板";
            return;
        }

        if (string.IsNullOrEmpty(OutputDirectory) || !Directory.Exists(OutputDirectory))
        {
            StatusMessage = "请选择有效的输出目录";
            return;
        }

        if (string.IsNullOrEmpty(OutputFileNameTemplate)) OutputFileNameTemplate = "{序号}_{时间}";

        // 验证数据源
        List<Dictionary<string, string>> dataRows;

        if (IsMultiTableMode)
        {
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

            // 使用合并后的数据
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

        // 校验身份证提取设置
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

            for (var i = 0; i < dataRows.Count; i++)
            {
                var rowData = new Dictionary<string, string>(dataRows[i]);

                // 添加一些特殊变量
                rowData["序号"] = (i + 1).ToString();
                rowData["时间"] = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                rowData["日期"] = DateTime.Now.ToString("yyyy-MM-dd");

                // 处理身份证信息提取 (保持原有逻辑)
                if (EnableIdCardExtraction && !string.IsNullOrEmpty(SelectedIdCardColumn) &&
                    rowData.TryGetValue(SelectedIdCardColumn, out var idCard) &&
                    !string.IsNullOrEmpty(idCard))
                    try
                    {
                        // 在日志中输出原始身份证号，方便调试
                        Debug.WriteLine($"处理身份证号: {idCard}");

                        // 添加提取的信息
                        var gender = _idCardService.ExtractGender(idCard);
                        var birthDate = _idCardService.ExtractBirthDate(idCard);
                        var region = _idCardService.ExtractRegion(idCard);

                        // 输出提取结果到日志
                        Debug.WriteLine($"提取结果: 性别={gender}, 生日={birthDate}, 地区={region}");

                        rowData["身份证性别"] = gender;
                        rowData["身份证出生日期"] = birthDate;
                        rowData["身份证籍贯"] = region;

                        // 提供更多格式的出生日期
                        rowData["身份证出生年"] = _idCardService.ExtractBirthDate(idCard, "yyyy");
                        rowData["身份证出生月"] = _idCardService.ExtractBirthDate(idCard, "MM");
                        rowData["身份证出生日"] = _idCardService.ExtractBirthDate(idCard, "dd");
                        rowData["身份证年龄"] = CalculateAge(birthDate);
                    }
                    catch (Exception ex)
                    {
                        // 捕获异常但不中断处理，记录错误信息
                        Debug.WriteLine($"身份证信息提取出错: {ex.Message}");

                        // 设置默认值
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

                fileName = $"{fileName}.docx";
                var outputPath = Path.Combine(OutputDirectory, fileName);

                // 异步处理每个文档
                var progress =
                    new Progress<int>(value => { ProgressValue = (i * 100 + value) / _excelData.Rows.Count; });

                var result = await _wordService.ProcessTemplateAsync(WordTemplatePath, outputPath, rowData, progress);

                if (result.Success)
                {
                    successCount++;
                }
                else
                {
                    failCount++;
                    StatusMessage = $"处理第 {i + 1} 行数据时出错: {result.Message}";
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
}