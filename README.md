# wp_transform

`wp_transform` 是一个在 Windows 本机上运行的 Python 命令行工具，用来把 Word 文档导出为 PDF。

## 功能

- 输入 `.doc` 或 `.docx` 文件路径
- 输出对应的 `.pdf` 文件
- 默认导出到源文件同目录
- 支持通过参数指定输出路径
- 支持覆盖已有 PDF

## 原理

工具通过 `pywin32` 调用本机安装的 Microsoft Word，使用 `ExportAsFixedFormat` 导出 PDF。  
这种方式通常比纯解析库更能保留原始排版、字体、页眉页脚和分页效果。

## 环境要求

- Windows
- Python 3.10+
- 已安装 Microsoft Word

## 安装

方式一：

```bash
pip install -r requirements.txt
```

方式二：

```bash
pip install .
```

## 用法

直接运行脚本：

```bash
python main.py "D:\docs\example.docx"
```

安装后使用命令：

```bash
wp-transform "D:\docs\example.docx"
```

指定输出文件：

```bash
wp-transform "D:\docs\example.docx" -o "D:\output\example.pdf"
```

覆盖已有文件：

```bash
wp-transform "D:\docs\example.docx" --overwrite
```

## 示例输出

```text
Converted successfully: D:\docs\example.pdf
```

## 说明

- 如果没有安装 Microsoft Word，转换会失败。
- 这是一个适合个人电脑本地使用的方案，不适合无 Office 环境的服务器。

## GitHub 上传建议

如果你要上传到 GitHub，推荐至少执行以下命令：

```bash
git init
git add .
git commit -m "feat: add Word to PDF CLI tool"
```
