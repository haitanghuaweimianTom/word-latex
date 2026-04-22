# Word → LaTeX 命令行工具

将 Word 文档 (.docx) 精确转换为 LaTeX 代码，支持数学公式、表格和排版。

## 功能特性

- **段落和标题转换**: Word 标题 1-6 → LaTeX `\section`, `\subsection` 等
- **数学公式转换**: 支持 OMML/MathML 公式转换为 LaTeX
  - 希腊字母: α, β, γ, π, Σ 等
  - 分式、根号、上下标
  - 矩阵、大型运算符 (∑, ∫, ∏)
- **表格转换**: 转为 `tabular` 环境，支持多列表格
- **列表转换**: 有序/无序列表 → `enumerate`/`itemize` 环境
- **样式转换**: 加粗 `\textbf{}`, 斜体 `\textit{}`
- **图片提取**: 自动提取文档中的图片

## 安装

```bash
# 克隆项目
cd word2latex

# 创建虚拟环境
uv venv
source .venv/Scripts/activate  # Windows: .venv\Scripts\activate

# 安装依赖
uv pip install python-docx
```

## 使用方法

```bash
# 基本用法 - 输出到标准输出
python -m word2latex document.docx

# 输出到文件
python -m word2latex document.docx -o output.tex

# 指定图片目录
python -m word2latex document.docx --img-dir ./images

# 仅输出文档内容（不含导言区）
python -m word2latex document.docx --package-mode

# 显示详细信息
python -m word2latex document.docx -v
```

## 命令行选项

| 选项 | 说明 |
|------|------|
| `INPUT_FILE` | Word 文档路径 (.docx) |
| `-o, --output PATH` | 输出文件路径 |
| `--img-dir PATH` | 图片提取目录 (默认: ./figures) |
| `--package-mode` | 仅输出 Document Body |
| `-v, --verbose` | 输出转换详情 |

## 测试

```bash
# 运行单元测试
python -m pytest tests/test_converter.py -v

# 创建测试文档
python tests/create_samples.py

# 运行所有测试
python run_test.py
```

## 项目结构

```
word2latex/
├── SPEC.md                 # 规格说明书
├── README.md               # 本文件
├── src/
│   ├── __init__.py
│   ├── __main__.py         # 模块入口
│   ├── cli.py              # 命令行接口
│   ├── converter.py        # 主转换引擎
│   ├── mathml_to_latex.py  # 数学公式转换器
│   └── table_converter.py  # 表格转换器
└── tests/
    ├── create_samples.py   # 测试样例生成器
    ├── test_converter.py   # 单元测试
    ├── formula_test.docx
    ├── table_test.docx
    ├── complex_test.docx
    └── math_test.docx
```

## 输出示例

### 表格

```latex
\begin{tabular}{|c|c|c|c|}
\hline
姓名 & 年龄 & 城市 & 职业 \\
\hline
张三 & 28 & 北京 & 工程师 \\
\hline
李四 & 35 & 上海 & 设计师 \\
\hline
\end{tabular}
```

### 数学公式 (需要 OMML/MathML 支持)

```latex
% 内联公式
二次方程: \(x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}\)

% 行间公式
\[\int_0^\infty e^{-x} dx = 1\]
```

## 已知限制

- 数学公式转换需要 Word 文档中嵌入 OMML/MathML
- 复杂表格（如跨行跨列）支持有限
- 某些特殊样式可能无法完全保留

## 许可证

MIT