using System.Diagnostics;
using System.IO;
using System.Text.Json;

namespace DocTransform.Models;

/// <summary>
///     应用程序设置类，用于保存和加载用户设置
/// </summary>
public class AppSettings
{
    /// <summary>
    ///     单例实例
    /// </summary>
    private static AppSettings _instance;

    /// <summary>
    ///     设置文件路径
    /// </summary>
    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "DocTransform",
        "settings.json");

    /// <summary>
    ///     获取设置实例
    /// </summary>
    public static AppSettings Instance => _instance ??= Load();

    /// <summary>
    ///     上次使用的输出目录
    /// </summary>
    public string LastOutputDirectory { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

    /// <summary>
    ///     从文件加载设置
    /// </summary>
    /// <returns>设置实例</returns>
    public static AppSettings Load()
    {
        try
        {
            // 确保设置目录存在
            var directory = Path.GetDirectoryName(SettingsPath);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);

            // 如果设置文件存在，则加载它
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);

                // 验证加载的输出目录是否存在，如果不存在则使用默认值
                if (settings != null && !string.IsNullOrEmpty(settings.LastOutputDirectory) &&
                    Directory.Exists(settings.LastOutputDirectory))
                    return settings;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"加载应用程序设置时出错: {ex.Message}");
        }

        // 返回默认设置
        return new AppSettings();
    }

    /// <summary>
    ///     保存设置到文件
    /// </summary>
    public void Save()
    {
        try
        {
            var directory = Path.GetDirectoryName(SettingsPath);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);

            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsPath, json);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"保存应用程序设置时出错: {ex.Message}");
        }
    }
}