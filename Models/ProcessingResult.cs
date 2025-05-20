namespace DocTransform.Models;

/// <summary>
///     处理结果信息
/// </summary>
public class ProcessingResult
{
    /// <summary>
    ///     是否成功
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    ///     结果消息
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    ///     生成的文件路径（如果成功）
    /// </summary>
    public string? FilePath { get; set; }

    /// <summary>
    ///     创建成功结果
    /// </summary>
    public static ProcessingResult Succeed(string message, string? filePath = null)
    {
        return new ProcessingResult { Success = true, Message = message, FilePath = filePath };
    }

    /// <summary>
    ///     创建失败结果
    /// </summary>
    public static ProcessingResult Fail(string message)
    {
        return new ProcessingResult { Success = false, Message = message };
    }
}