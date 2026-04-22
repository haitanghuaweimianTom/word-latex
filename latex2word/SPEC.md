# LaTeX → Word 命令行工具规格说明书

## 1. 项目概述

**项目名称**: latex2word
**项目类型**: Python 命令行工具
**核心功能**: 将 LaTeX 代码转换为 Word 文档 (.docx)，支持数学公式、表格和排版。
**目标用户**: 学术研究者、技术写作者、教育工作者。

---

## 2. 功能规格

### 2.1 核心功能

- **段落文本转换**: LaTeX 段落 → Word 段落
- **标题层级转换**: `\section`, `\subsection` 等 → Word 标题样式
- **数学公式转换**:
  - 行内公式 `\(` `\)` → Word 内联公式
  - 行间公式 `\[` `\]` 或 `$$` `$$` → Word 显示公式
  - 支持: 分式、根号、上下标、矩阵、希腊字母
- **表格转换**: `tabular`/`table` 环境 → Word 表格
- **列表转换**: `enumerate` → 有序列表，`itemize` → 无序列表
- **加粗/斜体**: `\textbf{}`, `\textit{}` → Word 加粗/斜体
- **图片引用**: `\includegraphics` → 保留图片路径注释

### 2.2 命令行接口

```
python -m latex2word [OPTIONS] INPUT_FILE

参数:
  INPUT_FILE              LaTeX 文件路径 (.tex)
  -o, --output PATH       输出文件路径（默认: stdout）
  --docx-format           输出为 .docx（默认）
  --package-mode          仅转换 Document Body（不含导言区）
  -v, --verbose           输出转换详情
  --version              显示版本
  --help                  显示帮助
```

---

## 3. 技术方案

- **语言**: Python 3.10+
- **Word 生成**: `python-docx` 库创建 .docx
- **LaTeX 解析**: 正则表达式 + 递归解析
- **数学公式处理**: 使用 OMML (Office Math Markup Language)
- **CLI 框架**: 标准库 `argparse`

---

## 4. 目录结构

```
latex2word/
├── SPEC.md
├── README.md
├── src/
│   ├── __init__.py
│   ├── __main__.py
│   ├── parser.py          # LaTeX 解析器
│   ├── converter.py       # 转换引擎
│   ├── math_converter.py   # 数学公式转换
│   ├── table_converter.py  # 表格转换
│   └── cli.py             # CLI 接口
└── tests/
    ├── test_parser.py
    ├── test_converter.py
    └── sample_files/
```

---

## 5. 验收标准

- ✅ 运行 `python -m latex2word sample.tex` 成功输出 .docx 文件
- ✅ 段落、标题正确转换为 Word 样式
- ✅ 数学公式（行内和行间）正确转换
- ✅ 表格正确转换为 Word 表格
- ✅ 有序/无序列表正确转换
- ✅ 加粗、斜体样式正确应用
- ✅ 错误输入给出友好提示