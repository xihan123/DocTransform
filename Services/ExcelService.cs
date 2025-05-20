using System.IO;
using ClosedXML.Excel;
using DocTransform.Models;

namespace DocTransform.Services;

/// <summary>
///     Excel文件处理服务
/// </summary>
public class ExcelService
{
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
}