# Word-LaTeX 双向转换工具

一个完整的 Word 与 LaTeX 文档相互转换工具集，支持命令行和 Web 界面两种使用方式。

## 功能特性

- **Word → LaTeX**: 将 .docx 文档转换为可编译的 LaTeX 代码
  - 支持数学公式（行内/行间）
  - 支持表格、图片提取
  - 支持标题、列表、加粗斜体等格式

- **LaTeX → Word**: 将 .tex 文件转换为 .docx 文档
  - 支持 LaTeX 数学公式
  - 支持表格转换
  - 保留文档结构

- **Web 界面**: 提供可视化的双向上传转换服务

---

## 项目结构

```
word-latex/
├── docx-tools/          # Web 服务（Flask）
│   ├── app.py          # 主程序
│   └── requirements.txt
├── word2latex/          # Word → LaTeX 转换工具
│   ├── src/
│   │   ├── converter.py       # 主转换引擎
│   │   ├── mathml_to_latex.py # 数学公式转换
│   │   └── table_converter.py # 表格转换
│   └── tests/
├── latex2word/          # LaTeX → Word 转换工具
│   ├── src/
│   │   ├── parser.py          # LaTeX 解析器
│   │   ├── converter.py       # 转换引擎
│   │   └── math_converter.py   # 数学公式转换
│   └── tests/
└── figures/             # 图片提取目录
```

---

## 安装

### 方式一：Web 界面（推荐）

```bash
# 进入 docx-tools 目录
cd docx-tools

# 安装依赖
pip install -r requirements.txt
```

### 方式二：命令行工具

```bash
# Word → LaTeX 依赖
pip install python-docx

# LaTeX → Word 依赖
pip install python-docx
```

---

## 使用方法

### 方法一：Web 界面

```bash
cd docx-tools
python app.py
```

启动后访问 **http://127.0.0.1:5000**

![界面预览](https://img.shields.io/badge/Web-5000-orange)

### 方法二：命令行

#### Word → LaTeX

```bash
cd word2latex

# 基本用法（输出到标准输出）
python -m word2latex input.docx

# 输出到文件
python -m word2latex input.docx -o output.tex

# 仅输出正文（不含导言区）
python -m word2latex input.docx --package-mode

# 指定图片目录
python -m word2latex input.docx --img-dir ./my-figures

# 详细模式
python -m word2latex input.docx -v
```

**输出示例**：
```latex
\documentclass{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{amsmath,amssymb}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{geometry}
\usepackage{enumerate}

\begin{document}

\section{标题}

正文内容\ldots

\end{document}
```

#### LaTeX → Word

```bash
cd latex2word

# 基本用法
python -m latex2word input.tex

# 输出到指定文件
python -m latex2word input.tex -o output.docx

# 详细模式
python -m latex2word input.tex -v
```

---

## 命令行参数详解

### word2latex

| 参数 | 说明 |
|------|------|
| `INPUT_FILE` | 输入的 Word 文档路径 (.docx) |
| `-o, --output` | 输出文件路径（默认输出到 stdout） |
| `--img-dir` | 图片提取目录（默认: ./figures） |
| `--package-mode` | 仅输出 Document Body（不含导言区） |
| `-v, --verbose` | 输出转换详情 |
| `--version` | 显示版本 |
| `--help` | 显示帮助信息 |

### latex2word

| 参数 | 说明 |
|------|------|
| `INPUT_FILE` | 输入的 LaTeX 文件路径 (.tex) |
| `-o, --output` | 输出文件路径（默认与输入同名） |
| `--package-mode` | 仅转换 Document Body（不含导言区） |
| `-v, --verbose` | 输出转换详情 |
| `--version` | 显示版本 |
| `--help` | 显示帮助信息 |

---

## 示例

### 示例 1：学术论文转换

将 Word 论文转换为 LaTeX，直接用于 Overleaf：

```bash
cd word2latex
python -m word2latex paper.docx -o paper.tex
```

然后在 Overleaf 中上传 `paper.tex` 即可编译。

### 示例 2：LaTeX 转 Word

将 LaTeX 论文转换为 Word，方便编辑：

```bash
cd latex2word
python -m latex2word paper.tex -o paper.docx
```

### 示例 3：Web 界面转换

1. 启动服务：
```bash
cd docx-tools
python app.py
```

2. 打开浏览器访问 `http://127.0.0.1:5000`

3. 拖拽或选择文件进行转换

---

## 输出格式说明

### 完整文档模式（默认）

生成的 LaTeX 包含完整文档结构，可直接编译：

```latex
\documentclass{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{amsmath,amssymb}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{geometry}
\usepackage{enumerate}

\begin{document}

% 文档内容

\end{document}
```

### 包模式（--package-mode）

仅输出文档内容体，不含导言区，适合嵌入到已有 LaTeX 文档：

```latex
\section{标题}

正文内容...

\begin{table}
% 表格内容
\end{table}
```

---

## 依赖说明

### Python 版本

- Python 3.10+

### 主要依赖

- `python-docx` >= 1.0.0 - Word 文档读写
- `flask` >= 2.0.0 - Web 框架（仅 Web 界面需要）
- `werkzeug` - WSGI 工具（Web 界面依赖）

---

## 常见问题

### Q: 转换后公式显示异常？

A: 检查是否正确加载了 `amsmath` 包。完整模式下会自动加载，如使用包模式请确保手动添加：

```latex
\usepackage{amsmath,amssymb}
```

### Q: 图片没有提取？

A: 确保 Word 文档中的图片是嵌入的（不是链接），并检查输出目录下是否有图片文件。

### Q: Web 界面无法访问？

A: 确保端口 5000 未被占用，可修改 `app.py` 中的端口：

```python
app.run(host='0.0.0.0', port=8080)  # 改为 8080
```

---

## 许可证

MIT License

---

## 作者

haitanghuaweimianTom