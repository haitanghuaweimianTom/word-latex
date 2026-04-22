# Word → LaTeX 命令行工具规格说明书

## 1. 项目概述

**项目名称**: word2latex
**项目类型**: Python 命令行工具
**核心功能**: 将 Word 文档 (.docx) 精确转换为 LaTeX 代码，支持数学公式、图片、表格和排版。
**目标用户**: 学术研究者、技术写作者、教育工作者。

---

## 2. 功能规格

### 2.1 核心功能

- **段落文本转换**: 普通段落转为 `\par` 或空行分隔的 LaTeX 代码
- **标题层级转换**: Word 标题 1-6 → LaTeX `\section`, `\subsection`, `\subsubsection`, `\paragraph`, `\subparagraph`
- **数学公式转换**:
  - Word 线性公式（OMML MathML）→ LaTeX 内联/行间公式
  - 支持上下标、分式、根号、矩阵、希腊字母、运算符
  - `\frac{}{}`, `\sqrt{}`, `^{}`, `_{}`, `\begin{matrix}` 等
- **表格转换**: 转为 `tabular` 或 `longtable` 环境，支持多列表格
- **图片转换**: 提取图片为独立文件，生成 `\includegraphics{}` 命令
- **列表转换**: 有序列表 → `enumerate` 环境，无序列表 → `itemize` 环境
- **加粗/斜体**: `\textbf{}`, `\textit{}`，可组合
- **样式继承**: 识别并保留段落样式名称

### 2.2 命令行接口

```
python -m word2latex [OPTIONS] INPUT_FILE

参数:
  INPUT_FILE              Word 文档路径 (.docx)
  -o, --output PATH       输出文件路径（默认: stdout）
  --img-dir PATH          图片提取目录（默认: ./figures）
  --package-mode          仅输出 Document Body（不含导言区），默认: 完整文档
  -v, --verbose           输出转换详情
  --version              显示版本
  --help                  显示帮助
```

### 2.3 输出格式

- 完整文档: 包含 `\documentclass`, `\usepackage`, `\begin{document}`, `\end{document}`
- 纯正文模式 (`--package-mode`): 仅输出文档内容体（不含导言区）

---

## 3. 验收标准

- ✅ 运行 `python -m word2latex sample.docx` 成功输出 LaTeX 代码
- ✅ 段落、标题正确转换为对应 LaTeX 命令
- ✅ 数学公式（行内和行间）正确转换
- ✅ 表格正确转换为 `tabular` 环境
- ✅ 图片提取并生成 `\includegraphics` 命令
- ✅ 有序/无序列表正确转换为 list 环境
- ✅ 加粗、斜体样式正确应用
- ✅ 错误输入给出友好提示
- ✅ 包含测试样例和测试脚本

---

## 4. 技术方案

- **语言**: Python 3.10+
- **Word 解析**: `python-docx` 库读取 docx XML 结构
- **数学公式解析**: 直接解析 Word OOXML 的 MathML → LaTeX 转换规则表
- **图片处理**: `zipfile` + `os` 提取 docx 内部图片
- **CLI 框架**: 标准库 `argparse`
- **无外部 LaTeX 依赖**: 纯 Python 实现，输出纯文本 LaTeX 代码

---

## 5. 目录结构

```
word2latex/
├── SPEC.md
├── README.md
├── src/
│   ├── __init__.py
│   ├── __main__.py
│   ├── converter.py       # 主转换引擎
│   ├── mathml_to_latex.py # MathML → LaTeX 转换器
│   ├── table_converter.py # 表格转换器
│   └── cli.py             # CLI 接口
├── tests/
│   ├── sample_formulas.py # 公式测试样例生成
│   ├── sample_complex.py  # 复杂文档测试样例
│   └── test_converter.py  # 单元测试
└── run_test.py            # 一键运行所有测试
```