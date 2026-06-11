![DocTransform](https://socialify.git.ci/xihan123/DocTransform/image?description=1&forks=1&issues=1&language=1&name=1&owner=1&pulls=1&stargazers=1&theme=Auto)

![.NET](https://img.shields.io/badge/.NET-9.0-blue.svg)
![WPF](https://img.shields.io/badge/UI-WPF-5C2D91.svg)
![License](https://img.shields.io/badge/License-GPL--3.0-orange.svg)
![版本](https://img.shields.io/github/v/release/xihan123/DocTransform)
![下载量](https://img.shields.io/github/downloads/xihan123/DocTransform/total)
[![.NET Desktop](https://github.com/xihan123/DocTransform/actions/workflows/dotnet-desktop.yml/badge.svg)](https://github.com/xihan123/DocTransform/actions/workflows/dotnet-desktop.yml)

# DocTransform

Excel 到 Word/Excel 模板的批量文档生成工具。

DocTransform 读取 `.xlsx` 数据源，将每一行数据填入 Word 或 Excel 模板中的 `{字段名}` 占位符，
批量生成合同、证书、通知书、登记表等格式固定但内容不同的文档。

## 目录

- [功能特性](#功能特性)
- [环境要求](#环境要求)
- [项目结构](#项目结构)
- [安装](#安装)
- [使用说明](#使用说明)
  - [准备数据源](#准备数据源)
  - [准备模板](#准备模板)
  - [生成文档](#生成文档)
  - [多表格模式](#多表格模式)
  - [图片替换](#图片替换)
  - [身份证信息提取](#身份证信息提取)
- [验证与构建](#验证与构建)
- [常见问题](#常见问题)
- [许可证](#许可证)

## 功能特性

- **Excel 数据源**：读取 `.xlsx` 第一行表头和后续数据行，保留日期等常见单元格显示格式。
- **Word 模板生成**：替换正文、表格、页眉、页脚中的 `{字段名}` 占位符，并尽量保留原有文字格式。
- **Excel 模板生成**：替换工作簿中各工作表的 `{字段名}` 占位符，保留单元格样式。
- **批量输出**：每条数据生成一份 Word、Excel，或同时生成两种文件。
- **多表格模式**：导入多个工作表或多个 Excel 文件，按公共键列合并后再生成文档。
- **输出文件名模板**：使用 `{姓名}_{序号}_{时间}` 这类模板生成文件名，并自动清理非法字符。
- **图片替换**：把图片目录名作为 Excel 模板占位符，例如 `{头像}`，按指定列匹配图片文件。
- **身份证字段**：可从身份证号列提取性别、出生日期、年龄、籍贯等常用字段。
- **拖放导入**：支持把 Excel 数据源或 Word 模板拖到主窗口。

## 环境要求

- Windows 10/11。
- .NET 9 Desktop Runtime，或使用 Release 中已按当前发布配置打包的版本。
- 可读写数据源、模板和输出目录的文件权限。

开发环境：

- .NET SDK 9.0，仓库 `global.json` 当前指定 `9.0.203`。
- Visual Studio 2022 或 Rider，需支持 .NET 9 WPF 项目。

## 安装

从 Release 下载：

1. 打开 [Releases](https://github.com/xihan123/DocTransform/releases)。
2. 下载对应架构的压缩包或可执行文件。
3. 解压后运行 `DocTransform-x64.exe` 或对应版本。

源码运行：

```powershell
dotnet restore DocTransform.sln
dotnet build DocTransform.sln -c Debug
dotnet run --project DocTransform.csproj
```

## 使用说明

### 准备数据源

数据源使用 `.xlsx` 文件。第一行作为列名，后续每一行作为一条记录。

示例：

```text
姓名 | 身份证号 | 部门 | 证书编号
张三 | 110101199001011234 | 技术部 | A-001
李四 | 110101199202022345 | 综合部 | A-002
```

列名会直接对应模板占位符。上面的 `姓名` 列对应 `{姓名}`。

### 准备模板

Word 和 Excel 模板都使用单层花括号占位符：

```text
姓名：{姓名}
部门：{部门}
证书编号：{证书编号}
生成日期：{时间}
```

内置可用字段：

```text
{序号}
{时间}
{日期}
```

启用身份证信息提取后，还可以在模板中使用：

```text
{身份证性别}
{身份证出生日期}
{身份证籍贯}
{身份证年龄}
{身份证出生年}
{身份证出生月}
{身份证出生日}
```

### 生成文档

1. 选择 Excel 数据源。
2. 选择 Word 模板，或启用并选择 Excel 模板。
3. 选择输出目录。
4. 设置输出文件名模板，例如 `{序号}_{姓名}_{时间}`。
5. 点击 `生成文档`。

如果同时选择 Word 模板和 Excel 模板，同一条数据会输出一份 `.docx` 和一份 `.xlsx`。

### 多表格模式

多表格模式用于把多个工作表或多个 Excel 文件合并成一份数据。

使用方式：

1. 打开 `多表格模式`。
2. 添加 Excel 文件。
3. 选择公共键列，例如 `身份证号`、`编号` 或 `姓名`。
4. 确认合并后的字段列表，再生成文档。

公共键列必须存在于参与合并的表中。合并后字段同样使用 `{字段名}` 写入模板。

### 图片替换

图片替换只作用于 Excel 模板。

使用方式：

1. 启用 Excel 模板。
2. 打开 `启用图片替换`。
3. 添加图片目录，例如目录名为 `头像`。
4. 在 Excel 模板中写入 `{头像}`。
5. 选择用于匹配图片文件名的数据列，例如 `姓名` 或 `证书编号`。

程序会扫描图片目录中的常见图片文件，并把匹配到的图片插入占位符所在单元格。
图片填充模式可选择适应或填充，并可调整填充比例。

### 身份证信息提取

启用后选择身份证号所在列。生成时会把身份证号解析为性别、出生日期、年龄、籍贯等字段。

身份证号解析失败时，对应字段会写入 `未知`，不会中断整批生成。

## 验证与构建

还原依赖：

```powershell
dotnet restore DocTransform.sln
```

Debug 构建：

```powershell
dotnet build DocTransform.sln -c Debug
```

Release 构建：

```powershell
dotnet build DocTransform.sln -c Release
```

发布 x64：

```powershell
dotnet publish DocTransform.csproj -c Release -r win-x64 --output ./publish-x64
```

发布 x86：

```powershell
dotnet publish DocTransform.csproj -c Release -r win-x86 --output ./publish-x86
```

## 常见问题

### Q: 占位符应该写成什么格式？

**A:** 使用单层花括号，例如 `{姓名}`。旧文档中的 `{{姓名}}`、`[[姓名]]` 不是当前实现使用的格式。

### Q: Word 模板中表格、页眉、页脚里的占位符会替换吗？

**A:** 会。程序会处理正文、表格、页眉和页脚中的占位符。

### Q: 可以只生成 Excel，不生成 Word 吗？

**A:** 可以。启用 Excel 模板并选择 `.xlsx` 模板即可。也可以同时选择 Word 模板和 Excel 模板。

### Q: 图片为什么没有插入？

**A:** 先检查三点：是否启用了 Excel 模板，模板中是否存在 `{目录名}` 占位符，图片目录是否设置了匹配列。

### Q: 输出文件名重复怎么办？

**A:** 建议在输出文件名模板中加入 `{序号}`、`{时间}` 或业务唯一字段，例如 `{证书编号}`。

## 许可证

本项目基于 [GNU General Public License v3.0](LICENSE) 发布。

---
