# Academic PDF Renamer

推荐 GitHub 仓库名：`paper-pdf-renamer`

一个用于批量整理学术 PDF 文件名的 Windows 桌面小工具。程序会读取 PDF 前几页文本，自动识别 DOI，并通过 CrossRef 获取论文标题、期刊、年份、接受日期、出版日期、第一作者等信息，然后按自定义规则生成新的文件名。

## 功能特点

- 批量扫描指定文件夹下的 PDF 文件
- 自动提取 DOI，并通过 CrossRef 获取论文元数据
- 支持标题、期刊名、期刊缩写、年份、接受日期、出版日期、第一作者、DOI 等命名字段
- 支持在图形界面中自由调整字段顺序和分隔符
- 重命名前可预览新文件名
- 自动避免覆盖同名文件
- 配置会保存到用户目录下的 `.pdf_renamer_config.json`

## 运行方式

### 方式一：直接运行打包好的程序

双击：

```text
PDF_Renamer.exe
```

### 方式二：用 Python 运行源码

先安装依赖：

```bash
pip install pypdf
```

然后运行：

```bash
python pdf_renamer.py
```

也可以直接双击：

```text
启动.bat
```

## 打包为 exe

在项目目录中双击：

```text
build.bat
```

或在命令行运行：

```bash
python -m pip install --upgrade pypdf pyinstaller
pyinstaller --onefile --windowed --noconfirm --clean --name PDF_Renamer pdf_renamer.py
```

打包完成后，生成的程序会复制到项目根目录：

```text
PDF_Renamer.exe
```

## 使用步骤

1. 打开程序。
2. 点击“选择文件夹”，选择存放 PDF 文献的文件夹。
3. 在“命名规则”区域选择需要的字段和分隔符。
4. 点击“扫描并提取”，等待程序读取元数据。
5. 检查“新文件名（预览）”是否符合预期。
6. 点击“执行重命名”。

## 建议

`PDF_Renamer.exe` 是打包后的成品文件，体积较大。如果只是保存源码，建议不要提交 exe 文件；如果想让别人直接下载使用，可以在 GitHub 的 Releases 页面上传 exe。
