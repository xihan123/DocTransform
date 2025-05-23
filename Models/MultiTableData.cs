using System.Collections.ObjectModel;

namespace DocTransform.Models;

public class MultiTableData
{
    public ObservableCollection<ExcelData> Tables { get; set; } = new();

    // 合并后的所有列（无重复）
    public List<string> AllHeaders
    {
        get
        {
            var uniqueHeaders = new HashSet<string>();
            foreach (var table in Tables)
            foreach (var header in table.Headers)
                uniqueHeaders.Add(header);

            return new List<string>(uniqueHeaders);
        }
    }

    // 在所有表中共同存在的列
    public List<string> CommonHeaders
    {
        get
        {
            if (Tables.Count == 0) return new List<string>();

            // 以第一个表的列为基础
            var common = new HashSet<string>(Tables[0].Headers);

            // 与其他表求交集
            for (var i = 1; i < Tables.Count; i++) common.IntersectWith(Tables[i].Headers);

            return new List<string>(common);
        }
    }

    // 合并后的数据记录
    public List<Dictionary<string, string>> MergedRows { get; private set; } = new();

    // 选中的用于匹配记录的键列（如身份证号、姓名等）
    public string KeyColumn { get; set; } = string.Empty;

    public int TotalRowCount
    {
        get
        {
            var count = 0;
            foreach (var table in Tables) count += table.RowCount;
            return count;
        }
    }

    // 合并数据
    public void MergeData(string keyColumn)
    {
        if (string.IsNullOrEmpty(keyColumn) || Tables.Count == 0) return;

        KeyColumn = keyColumn;
        MergedRows.Clear();

        // 使用字典存储合并后的数据，键是keyColumn的值
        var mergedData = new Dictionary<string, Dictionary<string, string>>();

        // 处理每个表格
        foreach (var table in Tables)
        foreach (var row in table.Rows)
        {
            // 确保行包含键列
            if (!row.TryGetValue(keyColumn, out var keyValue) || string.IsNullOrEmpty(keyValue)) continue;

            // 如果这是一个新的键值，创建新行
            if (!mergedData.TryGetValue(keyValue, out var mergedRow))
            {
                mergedRow = new Dictionary<string, string>();
                mergedData[keyValue] = mergedRow;
            }

            // 合并这一行的所有数据到合并行
            foreach (var pair in row)
                // 如果值非空，或者合并行中不存在此列，则添加/更新
                if (!string.IsNullOrEmpty(pair.Value) || !mergedRow.ContainsKey(pair.Key))
                    mergedRow[pair.Key] = pair.Value;
        }

        // 将合并结果转换为列表
        MergedRows = new List<Dictionary<string, string>>(mergedData.Values);
    }

    public void Clear()
    {
        Tables.Clear();
        MergedRows.Clear();
        KeyColumn = string.Empty;
    }
}