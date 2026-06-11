using System.IO;
using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using DocTransform.Models;
using OfficeOpenXml;

namespace DocTransform.Services;

/// <summary>
///     Excel文件处理服务
/// </summary>
public class ExcelService
{
    private static readonly CultureInfo DisplayCulture = CultureInfo.GetCultureInfo("zh-CN");
    private static readonly IReadOnlyDictionary<int, string> BuiltInDateFormats = new Dictionary<int, string>
    {
        [14] = "m/d/yy",
        [15] = "d-mmm-yy",
        [16] = "d-mmm",
        [17] = "mmm-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "m/d/yy h:mm",
        [27] = "yyyy\"年\"m\"月\"d\"日\"",
        [30] = "m\"月\"d\"日\"",
        [31] = "yyyy\"年\"m\"月\"d\"日\"",
        [36] = "yyyy\"年\"m\"月\"d\"日\"",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mm:ss.0",
        [50] = "yyyy\"年\"m\"月\"d\"日\"",
        [57] = "yyyy\"年\"m\"月\"",
        [58] = "m\"月\"d\"日\""
    };

    // 静态构造函数，用于设置EPPlus的LicenseContext
    static ExcelService()
    {
        // 设置EPPlus的LicenseContext为非商业用途
        ExcelPackage.License.SetNonCommercialOrganization("xihan123");
    }

    /// <summary>
    ///     异步读取Excel文件
    /// </summary>
    /// <param name="filePath">Excel文件路径</param>
    /// <returns>提取的Excel数据</returns>
    public async Task<ExcelData> ReadExcelFileAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            var excelData = new ExcelData();

            // 验证文件存在
            if (!File.Exists(filePath)) throw new FileNotFoundException("Excel文件不存在", filePath);

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.First();

            // 获取标题行（第一行）
            var headerRow = worksheet.Row(1);
            var columnCount = headerRow.CellsUsed().Count();

            // 如果没有数据，抛出异常
            if (columnCount == 0) throw new InvalidOperationException("Excel文件不包含任何数据");

            for (var i = 1; i <= columnCount; i++)
            {
                var headerText = headerRow.Cell(i).GetString().Trim();
                if (!string.IsNullOrEmpty(headerText)) excelData.Headers.Add(headerText);
            }

            // 读取数据行
            var lastRow = worksheet.LastRowUsed().RowNumber();
            for (var rowNumber = 2; rowNumber <= lastRow; rowNumber++)
            {
                var dataRow = new Dictionary<string, string>();
                var row = worksheet.Row(rowNumber);

                for (var colIndex = 0; colIndex < excelData.Headers.Count; colIndex++)
                {
                    var header = excelData.Headers[colIndex];
                    var cellValue = GetCellDisplayText(row.Cell(colIndex + 1));
                    dataRow[header] = cellValue;
                }

                excelData.Rows.Add(dataRow);
            }

            return excelData;
        });
    }


    // 读取Excel文件中所有工作表的方法
    public async Task<List<ExcelData>> ReadAllSheetsAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            var result = new List<ExcelData>();

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        // 跳过空工作表
                        if (worksheet.Dimension == null) continue;

                        var sheetData = new ExcelData
                        {
                            SourceFileName = $"{Path.GetFileName(filePath)} - {worksheet.Name}"
                        };

                        // 读取列标题
                        var colCount = worksheet.Dimension.End.Column;
                        for (var col = 1; col <= colCount; col++)
                        {
                            var headerCell = worksheet.Cells[1, col].Text.Trim();
                            if (!string.IsNullOrEmpty(headerCell)) sheetData.Headers.Add(headerCell);
                        }

                        // 读取数据行
                        var rowCount = worksheet.Dimension.End.Row;
                        for (var row = 2; row <= rowCount; row++) // 从第二行开始，跳过标题行
                        {
                            var dataRow = new Dictionary<string, string>();
                            var hasData = false;

                            for (var col = 1; col <= colCount; col++)
                            {
                                if (col > sheetData.Headers.Count) continue;

                                var header = sheetData.Headers[col - 1];
                                var cellValue = GetCellDisplayText(worksheet.Cells[row, col]).Trim();

                                dataRow[header] = cellValue;
                                if (!string.IsNullOrEmpty(cellValue)) hasData = true;
                            }

                            if (hasData) sheetData.Rows.Add(dataRow);
                        }

                        // 只添加非空的工作表
                        if (sheetData.Headers.Count > 0 && sheetData.Rows.Count > 0) result.Add(sheetData);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取Excel文件时出错: {ex.Message}", ex);
            }

            return result;
        });
    }

    private static string GetCellDisplayText(IXLCell cell)
    {
        if (cell.IsEmpty()) return string.Empty;

        if (TryFormatDateCell(cell, out var formattedDate)) return formattedDate;

        try
        {
            return cell.GetFormattedString(DisplayCulture);
        }
        catch
        {
            return cell.GetString();
        }
    }

    private static string GetCellDisplayText(ExcelRangeBase cell)
    {
        if (cell.Value == null && string.IsNullOrEmpty(cell.Formula)) return string.Empty;

        if (TryFormatDateCell(cell, out var formattedDate)) return formattedDate;

        try
        {
            return cell.Text;
        }
        catch
        {
            return Convert.ToString(cell.Value, DisplayCulture) ?? string.Empty;
        }
    }

    private static bool TryFormatDateCell(IXLCell cell, out string formattedDate)
    {
        formattedDate = string.Empty;

        var format = GetNumberFormat(
            cell.Style.NumberFormat.Format,
            cell.Style.NumberFormat.NumberFormatId);

        if (!TryConvertExcelDateFormat(format, out var dotNetFormat)) return false;
        if (!TryGetCellDateTime(cell, out var dateTime)) return false;

        return TryFormatDateTime(dateTime, dotNetFormat, out formattedDate);
    }

    private static bool TryFormatDateCell(ExcelRangeBase cell, out string formattedDate)
    {
        formattedDate = string.Empty;

        var format = GetNumberFormat(
            cell.Style.Numberformat.Format,
            cell.Style.Numberformat.NumFmtID);

        if (!TryConvertExcelDateFormat(format, out var dotNetFormat)) return false;
        if (!TryGetCellDateTime(cell.Value, cell.Worksheet.Workbook.Date1904, out var dateTime)) return false;

        return TryFormatDateTime(dateTime, dotNetFormat, out formattedDate);
    }

    private static string? GetNumberFormat(string? format, int numberFormatId)
    {
        if (!string.IsNullOrWhiteSpace(format) &&
            !string.Equals(format, "General", StringComparison.OrdinalIgnoreCase))
            return format;

        return BuiltInDateFormats.TryGetValue(numberFormatId, out var builtInFormat)
            ? builtInFormat
            : null;
    }

    private static bool TryGetCellDateTime(IXLCell cell, out DateTime dateTime)
    {
        dateTime = default;

        try
        {
            var value = cell.Value;

            if (value.IsDateTime)
            {
                dateTime = value.GetDateTime();
                return true;
            }

            if (value.IsNumber)
                return TryConvertExcelSerialDate(
                    value.GetNumber(),
                    cell.Worksheet.Workbook.Use1904DateSystem,
                    out dateTime);

            if (value.IsText &&
                DateTime.TryParse(value.GetText(), DisplayCulture, DateTimeStyles.None, out dateTime))
                return true;
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static bool TryGetCellDateTime(object? value, bool use1904DateSystem, out DateTime dateTime)
    {
        dateTime = default;

        switch (value)
        {
            case DateTime dateTimeValue:
                dateTime = dateTimeValue;
                return true;
            case string text when DateTime.TryParse(text, DisplayCulture, DateTimeStyles.None, out dateTime):
                return true;
            case IConvertible convertible when value is not bool:
                try
                {
                    return TryConvertExcelSerialDate(
                        convertible.ToDouble(CultureInfo.InvariantCulture),
                        use1904DateSystem,
                        out dateTime);
                }
                catch
                {
                    return false;
                }
            default:
                return false;
        }
    }

    private static bool TryConvertExcelSerialDate(
        double serialDate,
        bool use1904DateSystem,
        out DateTime dateTime)
    {
        dateTime = default;

        if (double.IsNaN(serialDate) || double.IsInfinity(serialDate)) return false;

        try
        {
            dateTime = use1904DateSystem
                ? new DateTime(1904, 1, 1).AddDays(serialDate)
                : DateTime.FromOADate(serialDate);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryFormatDateTime(DateTime dateTime, string dotNetFormat, out string formattedDate)
    {
        formattedDate = string.Empty;

        try
        {
            formattedDate = dateTime.ToString(dotNetFormat, DisplayCulture);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryConvertExcelDateFormat(string? excelFormat, out string dotNetFormat)
    {
        dotNetFormat = string.Empty;

        var section = GetFirstFormatSection(excelFormat);
        if (string.IsNullOrWhiteSpace(section) ||
            string.Equals(section.Trim(), "General", StringComparison.OrdinalIgnoreCase))
            return false;

        var converted = new StringBuilder();
        var hasDateToken = false;
        char? previousToken = null;

        for (var i = 0; i < section.Length;)
        {
            var current = section[i];

            if (current == '"')
            {
                var endQuote = section.IndexOf('"', i + 1);
                if (endQuote < 0)
                {
                    AppendLiteral(converted, section[(i + 1)..]);
                    break;
                }

                AppendLiteral(converted, section.Substring(i + 1, endQuote - i - 1));
                i = endQuote + 1;
                continue;
            }

            if (current == '\\')
            {
                if (i + 1 < section.Length) AppendLiteral(converted, section[i + 1].ToString());
                i += 2;
                continue;
            }

            if (current is '_' or '*')
            {
                i += 2;
                continue;
            }

            if (current == '[')
            {
                var endBracket = section.IndexOf(']', i + 1);
                if (endBracket < 0) return false;

                i = endBracket + 1;
                continue;
            }

            if (IsTokenAt(section, i, "AM/PM", out var tokenLength) ||
                IsTokenAt(section, i, "A/P", out tokenLength))
            {
                converted.Append("tt");
                i += tokenLength;
                continue;
            }

            var token = char.ToLowerInvariant(current);
            if (token is 'y' or 'm' or 'd' or 'h' or 's')
            {
                var tokenStart = i;
                while (i < section.Length && char.ToLowerInvariant(section[i]) == token) i++;

                var length = i - tokenStart;
                hasDateToken = true;

                switch (token)
                {
                    case 'y':
                    case 'd':
                    case 's':
                        converted.Append(new string(token, length));
                        break;
                    case 'h':
                        converted.Append(new string('H', length));
                        break;
                    case 'm':
                        var nextToken = GetNextDateToken(section, i);
                        var isMinute = previousToken == 'h' || nextToken == 's';
                        converted.Append(new string(isMinute ? 'm' : 'M', length));
                        break;
                }

                previousToken = token;
                continue;
            }

            AppendLiteral(converted, current.ToString());
            i++;
        }

        if (!hasDateToken) return false;

        dotNetFormat = converted.ToString();
        if (dotNetFormat.Length == 1 && "yMdHhms".Contains(dotNetFormat, StringComparison.Ordinal))
            dotNetFormat = $"%{dotNetFormat}";

        return !string.IsNullOrEmpty(dotNetFormat);
    }

    private static string? GetFirstFormatSection(string? excelFormat)
    {
        if (string.IsNullOrWhiteSpace(excelFormat)) return null;

        var inQuote = false;

        for (var i = 0; i < excelFormat.Length; i++)
        {
            if (excelFormat[i] == '"') inQuote = !inQuote;
            if (excelFormat[i] == ';' && !inQuote) return excelFormat[..i];
        }

        return excelFormat;
    }

    private static char? GetNextDateToken(string format, int startIndex)
    {
        for (var i = startIndex; i < format.Length;)
        {
            var current = format[i];

            if (current == '"')
            {
                var endQuote = format.IndexOf('"', i + 1);
                if (endQuote < 0) return null;

                i = endQuote + 1;
                continue;
            }

            if (current is '\\' or '_' or '*')
            {
                i += 2;
                continue;
            }

            if (current == '[')
            {
                var endBracket = format.IndexOf(']', i + 1);
                if (endBracket < 0) return null;

                i = endBracket + 1;
                continue;
            }

            if (IsTokenAt(format, i, "AM/PM", out var tokenLength) ||
                IsTokenAt(format, i, "A/P", out tokenLength))
            {
                i += tokenLength;
                continue;
            }

            var token = char.ToLowerInvariant(current);
            if (token is 'y' or 'm' or 'd' or 'h' or 's') return token;

            i++;
        }

        return null;
    }

    private static bool IsTokenAt(string value, int startIndex, string token, out int tokenLength)
    {
        tokenLength = token.Length;

        return startIndex + token.Length <= value.Length &&
               string.Compare(value, startIndex, token, 0, token.Length, true, DisplayCulture) == 0;
    }

    private static void AppendLiteral(StringBuilder builder, string value)
    {
        if (string.IsNullOrEmpty(value)) return;

        builder.Append('\'');
        builder.Append(value.Replace("'", "''"));
        builder.Append('\'');
    }
}
