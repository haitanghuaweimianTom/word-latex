# LaTeX → Word 命令行工具

将 LaTeX 文档 (.tex) 转换为 Word 文档 (.docx)，支持数学公式、表格和排版。

## 功能特性

- **段落和标题转换**: LaTeX `\section`, `\subsection` → Word 标题样式
- **数学公式转换**:
  - LaTeX 公式 → 简化文本格式（因 Word OMML 复杂性限制）
  - 支持: 分式、根号、上下标、希腊字母
- **表格转换**: LaTeX `tabular` 环境 → Word 表格
- **列表转换**: LaTeX `enumerate`/`itemize` → Word 有序/无序列表
- **样式转换**: `\textbf{}` → 加粗, `\textit{}` → 斜体

## 安装

```bash
# 克隆项目
cd latex2word

# 复用 word2latex 的虚拟环境
source ../word2latex/.venv/Scripts/activate  # Windows: .venv\Scripts\activate

# 或者创建新环境
uv venv
source .venv/Scripts/activate
uv pip install python-docx
```

## 使用方法

```bash
# 基本用法 - 输出到 docx 文件
python -m latex2word document.tex

# 指定输出文件
python -m latex2word document.tex -o output.docx

# 显示详细信息
python -m latex2word document.tex -v
```

## 命令行选项

| 选项 | 说明 |
|------|------|
| `INPUT_FILE` | LaTeX 文件路径 (.tex) |
| `-o, --output PATH` | 输出文件路径 |
| `--package-mode` | 仅转换 Document Body |
| `-v, --verbose` | 输出转换详情 |

## 测试

```bash
# 运行单元测试
python -m pytest tests/test_converter.py -v

# 测试文档转换
python -m latex2word tests/sample_latex.tex -o tests/output.docx -v
```

## 项目结构

```
latex2word/
├── SPEC.md                 # 规格说明书
├── README.md               # 本文件
├── src/
│   ├── __init__.py
│   ├── __main__.py         # 模块入口
│   ├── cli.py              # 命令行接口
│   ├── parser.py           # LaTeX 解析器
│   ├── converter.py         # 转换引擎
│   ├── math_converter.py   # 数学公式转换器
│   └── table_converter.py  # 表格转换器
└── tests/
    ├── test_converter.py   # 单元测试
    ├── sample_latex.tex    # 测试样例
    └── output.docx         # 转换结果
```

## 已知限制

- 数学公式转换为简化文本格式（不是真正的 Word 公式对象）
- 不支持复杂的 LaTeX 宏和自定义命令
- 不支持 TikZ 图形和特殊包

## 许可证

MIT