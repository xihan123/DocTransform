using System.IO;
using ClosedXML.Excel;
using DocTransform.Models;
using OfficeOpenXml;

namespace DocTransform.Services;

/// <summary>
///     Excel文件处理服务
/// </summary>
public class ExcelService
{
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
                    var cellValue = row.Cell(colIndex + 1).GetString();
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
                                var cellValue = worksheet.Cells[row, col].Text.Trim();

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
}