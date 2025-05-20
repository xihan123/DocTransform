using System.Text.RegularExpressions;

namespace DocTransform.Services;

/// <summary>
///     身份证信息处理服务
/// </summary>
public class IdCardService
{
    // 中国地区代码字典
    private static readonly Dictionary<string, string> RegionCodeMap = new()
    {
        { "11", "北京" }, { "12", "天津" }, { "13", "河北" }, { "14", "山西" }, { "15", "内蒙古" },
        { "21", "辽宁" }, { "22", "吉林" }, { "23", "黑龙江" },
        { "31", "上海" }, { "32", "江苏" }, { "33", "浙江" }, { "34", "安徽" }, { "35", "福建" }, { "36", "江西" }, { "37", "山东" },
        { "41", "河南" }, { "42", "湖北" }, { "43", "湖南" }, { "44", "广东" }, { "45", "广西" }, { "46", "海南" },
        { "50", "重庆" }, { "51", "四川" }, { "52", "贵州" }, { "53", "云南" }, { "54", "西藏" },
        { "61", "陕西" }, { "62", "甘肃" }, { "63", "青海" }, { "64", "宁夏" }, { "65", "新疆" },
        { "71", "台湾" },
        { "81", "香港" }, { "82", "澳门" }
    };

    /// <summary>
    ///     检查身份证号码是否有效（基本格式检查）
    /// </summary>
    /// <param name="idCard">身份证号码</param>
    /// <returns>是否有效</returns>
    public bool IsValidIdCard(string idCard)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(idCard)) return false;

            // 过滤掉空格和其他非数字字符（可能由Excel格式导致）
            idCard = Regex.Replace(idCard, @"[^\dXx]", "");

            // 15位或18位
            if (idCard.Length != 15 && idCard.Length != 18) return false;

            // 检查数字格式
            if (idCard.Length == 15) return Regex.IsMatch(idCard, @"^\d{15}$");

            // 简化检查：18位身份证，前17位为数字，最后一位可以是数字或X
            return Regex.IsMatch(idCard, @"^\d{17}[\dXx]$");
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    ///     提取身份证中的性别信息
    /// </summary>
    /// <param name="idCard">身份证号码</param>
    /// <returns>性别：男/女</returns>
    public string ExtractGender(string idCard)
    {
        try
        {
            // 过滤掉空格和其他非数字字符
            idCard = Regex.Replace(idCard, @"[^\dXx]", "");

            if (idCard.Length != 15 && idCard.Length != 18) return "未知";

            int genderCode;
            if (idCard.Length == 15)
                genderCode = int.Parse(idCard.Substring(14, 1));
            else // 18位
                genderCode = int.Parse(idCard.Substring(16, 1));

            // 性别代码：奇数为男，偶数为女
            return genderCode % 2 == 0 ? "女" : "男";
        }
        catch
        {
            return "未知";
        }
    }

    /// <summary>
    ///     提取身份证中的出生日期
    /// </summary>
    /// <param name="idCard">身份证号码</param>
    /// <param name="format">日期格式</param>
    /// <returns>格式化的出生日期</returns>
    public string ExtractBirthDate(string idCard, string format = "yyyy-MM-dd")
    {
        try
        {
            // 过滤掉空格和其他非数字字符
            idCard = Regex.Replace(idCard, @"[^\dXx]", "");

            if (idCard.Length != 15 && idCard.Length != 18) return "无效日期";

            string birthDateStr;
            if (idCard.Length == 15)
            {
                birthDateStr = "19" + idCard.Substring(6, 6);

                // 解析为年月日
                var year = int.Parse(birthDateStr.Substring(0, 4));
                var month = int.Parse(birthDateStr.Substring(4, 2));
                var day = int.Parse(birthDateStr.Substring(6, 2));

                try
                {
                    return new DateTime(year, month, day).ToString(format);
                }
                catch
                {
                    return "无效日期";
                }
            }
            else
            {
                birthDateStr = idCard.Substring(6, 8);

                // 解析为年月日
                var year = int.Parse(birthDateStr.Substring(0, 4));
                var month = int.Parse(birthDateStr.Substring(4, 2));
                var day = int.Parse(birthDateStr.Substring(6, 2));

                try
                {
                    return new DateTime(year, month, day).ToString(format);
                }
                catch
                {
                    return "无效日期";
                }
            }
        }
        catch
        {
            return "无效日期";
        }
    }

    /// <summary>
    ///     提取身份证中的籍贯信息（精确到省/直辖市）
    /// </summary>
    /// <param name="idCard">身份证号码</param>
    /// <returns>籍贯</returns>
    public string ExtractRegion(string idCard)
    {
        try
        {
            // 过滤掉空格和其他非数字字符
            idCard = Regex.Replace(idCard, @"[^\dXx]", "");

            if (idCard.Length != 15 && idCard.Length != 18) return "未知地区";

            var provinceCode = idCard.Substring(0, 2);
            if (RegionCodeMap.TryGetValue(provinceCode, out var regionName)) return regionName;

            return "未知地区";
        }
        catch
        {
            return "未知地区";
        }
    }
}