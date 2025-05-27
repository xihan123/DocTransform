# DocTransform

![版本](https://img.shields.io/github/v/release/xihan123/DocTransform)
![语言](https://img.shields.io/badge/语言-C%23-brightgreen.svg)
![许可证](https://img.shields.io/badge/许可证-MIT-orange.svg)
![下载量](https://img.shields.io/github/downloads/xihan123/DocTransform/total)

## 📝 项目简介

DocTransform 是一款高效的 Excel 到 Word 数据映射应用程序，使用 C# 开发。它允许用户从 Excel 电子表格中提取数据，并将其映射到预定义的
Word 文档模板中，从而快速生成格式化的文档报告。

这个工具特别适合需要从 Excel 数据生成标准化 Word 文档的场景，如批量生成合同、报告、证书等文档处理任务。

## ✨ 主要功能

- **数据提取**: - 从 Excel 工作表中精确读取所需数据
- **灵活映射**:
    - 自动或手动将Excel中的数据列映射到Word模板中的占位符。
    - 支持多种数据匹配模式，例如：
        - **多表格模式**: 实现了多表格模式下的数据导入和匹配功能，支持多个Excel表格的数据导入、自动匹配、显示和操作，以及数据合并后的结果显示。
- **批量处理** - 一次性处理多条数据记录，生成多个文档
- **模板支持** - 使用预设 Word 、 Excel模板，保持输出文档的一致性
- **用户友好界面** - 简洁直观的操作流程
- **自定义配置** - 允许用户自定义映射规则和输出设置

## 🛠️ 技术栈

- **C#**: 主要编程语言。
- **WPF (Windows Presentation Foundation)**: 用于构建用户界面。
- **.NET**: 开发框架。

## 💻 系统要求

- Windows 操作系统
- .NET Framework 9.x
- 足够的系统权限来读写文件

## 📥 安装说明

### 直接下载

1. 从 [Releases](https://github.com/xihan123/DocTransform/releases) 页面下载最新版本的安装包
2. 解压下载的文件到您选择的目录
3. 运行 `DocTransform-x64.exe` 启动应用程序

## 📋 使用指南

### 基本用法

1. **准备数据和模板**
    - 创建或准备包含所需数据的 Excel 文件
    - 创建 Word 模板文档，在需要插入数据的位置添加占位符（例如 `{{字段名}}` 或 `[[字段名]]`）

2. **启动应用程序**
    - 运行 DocTransform.exe

3. **配置映射**
    - 选择源 Excel 文件
    - 选择目标 Word/Excel 模板
    - 设置输出目录
    - 配置字段映射关系
    - 图片映射是将目录名作为占位符(例如 `{{图片}}`)放在模板中，程序会自动将 Excel 中的图片数据插入到对应位置
    - 图片嵌入设置是指填充模式(如需填充头像，如2寸兔冠彩照大概是选择填充模式-填充比例65%左右,也和照片分辨率有关)

4. **生成文档**
    - 点击"生成文档"按钮
    - 等待处理完成
    - 查看生成的 Word/Excel 文档(如果同时输入了Word和Excel模板，生成的文档会在输出目录下生成两个文件)

### 高级用法

- **批量处理**：选择 Excel 中的多行数据进行批量文档生成
- **条件映射**：根据特定条件选择性地映射数据
- **保存配置**：保存映射配置以便下次使用
- **自定义输出**：配置输出文件的命名规则

## 🔍 常见问题

**Q: 如何处理特殊格式的 Excel 数据?**  
A: 在映射配置中，您可以指定数据格式转换规则。

**Q: 是否支持图片、表格等复杂元素的映射?**  
A: 是的，支持将 Excel 中的图片数据和表格结构映射到 Word 文档中。现在还支持将本地指定目录映射到模板中

**Q: 能够批量处理多个 Excel 文件吗?**  
A: ~~当前版本支持单次选择一个 Excel 文件，但可以处理文件中的多条数据记录。~~1.0.2+版本已支持批量处理多个 Excel 文件作为数据源。

## 🛠 故障排除

如果遇到问题，请尝试以下步骤：

1. 确保您使用的是最新版本的应用程序
2. 检查 Excel 和 Word 文件格式是否受支持
3. 验证占位符格式是否正确
4. 查看日志文件（位于应用程序目录下的 `logs` 文件夹）
5. 如果问题持续存在，请[提交 Issue](https://github.com/xihan123/DocTransform/issues/new)

## 🤝 如何贡献

我们欢迎任何形式的贡献！

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交您的更改 (`git commit -m '添加一些很棒的功能'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 打开一个 Pull Request

## 📜 许可证

该项目采用 MIT 许可证 - 详情请参见 [LICENSE](LICENSE) 文件

## 👨‍💻 作者

[xihan123](https://github.com/xihan123)

## 📊 项目状态

![GitHub stars](https://img.shields.io/github/stars/xihan123/DocTransform?style=social)
![GitHub forks](https://img.shields.io/github/forks/xihan123/DocTransform?style=social)
![GitHub issues](https://img.shields.io/github/issues/xihan123/DocTransform)
![GitHub pull requests](https://img.shields.io/github/issues-pr/xihan123/DocTransform)

---

如果您觉得这个工具有用，请考虑给项目点个 ⭐️！
