"""
LaTeX 解析器 - 解析 LaTeX 文档结构
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


@dataclass
class LaTeXElement:
    """LaTeX 元素基类"""
    pass


@dataclass
class TextElement(LaTeXElement):
    """文本元素"""
    content: str
    bold: bool = False
    italic: bool = False


@dataclass
class CommandElement(LaTeXElement):
    """命令元素"""
    name: str
    args: List[str] = field(default_factory=list)
    content: Optional[str] = None  # 命令后面的内容


@dataclass
class MathElement(LaTeXElement):
    """数学公式元素"""
    content: str
    display: bool = False  # 是否是行间公式


@dataclass
class EnvironmentElement(LaTeXElement):
    """环境元素"""
    name: str
    content: List[LaTeXElement] = field(default_factory=list)


@dataclass
class TableElement(LaTeXElement):
    """表格元素"""
    rows: List[List[str]] = field(default_factory=list)
    has_header: bool = True


@dataclass
class LaTeXDocument:
    """LaTeX 文档"""
    elements: List[LaTeXElement] = field(default_factory=list)
    preamble: str = ""


class LaTeXParser:
    """
    LaTeX 解析器

    支持解析:
    - 标题命令: \section, \subsection, etc.
    - 文本格式: \textbf, \textit, etc.
    - 数学公式: \(...\), \[...\], $...$, $$...$$
    - 表格: \begin{tabular}...\end{tabular}
    - 列表: \begin{itemize}, \begin{enumerate}
    """

    # LaTeX 命令模式
    COMMAND_PATTERN = r'\\([a-zA-Z]+|.)(?:\[[^\]]*\])?(\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\})*'

    # 标题命令
    HEADING_COMMANDS = {
        'section': 1,
        'subsection': 2,
        'subsubsection': 3,
        'paragraph': 4,
        'subparagraph': 5,
    }

    def __init__(self):
        self._pos = 0
        self._text = ""
        self._elements: List[LaTeXElement] = []

    def parse(self, tex: str) -> LaTeXDocument:
        """
        解析 LaTeX 文本

        Args:
            tex: LaTeX 源代码

        Returns:
            LaTeXDocument 对象
        """
        self._pos = 0
        self._text = tex
        self._elements = []

        # 提取导言区和正文
        doc_start = self._find_document_start(tex)
        if doc_start >= 0:
            preamble = tex[:doc_start]
            content = tex[doc_start:]
        else:
            preamble = ""
            content = tex

        # 解析正文
        self._parse_content(content)

        return LaTeXDocument(elements=self._elements, preamble=preamble)

    def _find_document_start(self, tex: str) -> int:
        """找到 \begin{document} 的位置"""
        match = re.search(r'\\begin\{document\}', tex)
        return match.start() if match else -1

    def _parse_content(self, content: str):
        """解析 LaTeX 内容"""
        # 移除 \end{document} 及其后的内容
        content = re.split(r'\\end\{document\}', content)[0]

        # 移除注释
        content = self._remove_comments(content)

        # 解析各种元素
        i = 0
        while i < len(content):
            matched = False

            # 跳过空白
            if content[i:i+2] == '\\%':
                i += 2
                continue

            # 跳过换行
            if content[i] == '\n':
                i += 1
                continue

            # 检查标题命令
            for cmd_name, level in self.HEADING_COMMANDS.items():
                pattern = rf'\\{cmd_name}\*?\{{'
                match = re.match(pattern, content[i:])
                if match:
                    title = self._extract_braced_content(content[i + match.end() - 1:])
                    self._elements.append(CommandElement(
                        name=cmd_name,
                        args=[title],
                        content=f"Heading {level}"
                    ))
                    i += match.end() + len(title) + 1
                    matched = True
                    break

            if matched:
                continue

            # 检查数学公式 (行间公式)
            if content[i:i+2] == '\\[':
                end = content.find('\\]', i)
                if end >= 0:
                    math_content = content[i+2:end]
                    self._elements.append(MathElement(content=math_content, display=True))
                    i = end + 2
                    matched = True

            if matched:
                continue

            # 检查数学公式 (行内公式 \( \))
            if content[i:i+2] == '\\(':
                end = content.find('\\)', i)
                if end >= 0:
                    math_content = content[i+2:end]
                    self._elements.append(MathElement(content=math_content, display=False))
                    i = end + 2
                    matched = True

            if matched:
                continue

            # 检查数学公式 (行内公式 $ $)
            if content[i] == '$':
                # 检查是否是 $$ (行间公式)
                if content[i:i+2] == '$$':
                    end = content.find('$$', i + 2)
                    if end >= 0:
                        math_content = content[i+2:end]
                        self._elements.append(MathElement(content=math_content, display=True))
                        i = end + 2
                        matched = True
                else:
                    end = content.find('$', i + 1)
                    if end >= 0:
                        math_content = content[i+1:end]
                        self._elements.append(MathElement(content=math_content, display=False))
                        i = end + 1
                        matched = True

            if matched:
                continue

            # 检查表格环境
            tabular_pattern = '\\begin{tabular}'
            if content[i:i+len(tabular_pattern)] == tabular_pattern:
                table_content = self._extract_environment(content, i, 'tabular')
                if table_content:
                    table = self._parse_table(table_content)
                    if table:
                        self._elements.append(table)
                    end_pattern = '\\end{tabular}'
                    i += len(table_content) + len(tabular_pattern) + len(end_pattern)
                    matched = True

            if matched:
                continue

            # 检查列表环境
            for list_env in ['itemize', 'enumerate']:
                begin_pattern = f'\\begin{{{list_env}}}'
                if content[i:i+len(begin_pattern)] == begin_pattern:
                    list_content = self._extract_environment(content, i, list_env)
                    if list_content:
                        items = self._parse_list(list_content, list_env)
                        self._elements.append(EnvironmentElement(name=list_env, content=items))
                        end_pattern = f'\\end{{{list_env}}}'
                        i += len(list_content) + len(begin_pattern) + len(end_pattern)
                        matched = True
                        break

            if matched:
                continue

            # 检查加粗命令
            bold_match = re.match(r'\\textbf\{([^}]*)\}', content[i:])
            if bold_match:
                self._elements.append(TextElement(content=bold_match.group(1), bold=True))
                i += len(bold_match.group(0))
                matched = True

            if matched:
                continue

            # 检查斜体命令
            italic_match = re.match(r'\\textit\{([^}]*)\}', content[i:])
            if italic_match:
                self._elements.append(TextElement(content=italic_match.group(1), italic=True))
                i += len(italic_match.group(0))
                matched = True

            if matched:
                continue

            # 检查其他命令
            cmd_match = re.match(r'\\([a-zA-Z]+)', content[i:])
            if cmd_match:
                i += len(cmd_match.group(0))
                continue

            # 检查反斜杠
            if content[i] == '\\':
                i += 1
                continue

            # 普通文本
            # 找到下一个特殊字符的位置
            next_special = len(content)
            for pattern in ['\\[', '\\(', '\\$', '$$', '\\textbf', '\\textit',
                           '\\begin{tabular}', '\\begin{itemize}', '\\begin{enumerate}']:
                idx = content.find(pattern, i)
                if idx >= 0 and idx < next_special:
                    next_special = idx

            if next_special > i:
                text = content[i:next_special]
                # 清理多余空白
                text = ' '.join(text.split())
                if text:
                    self._elements.append(TextElement(content=text))
                i = next_special
            else:
                i += 1

    def _remove_comments(self, content: str) -> str:
        """移除 LaTeX 注释"""
        lines = content.split('\n')
        result = []
        for line in lines:
            # 移除 % 后的注释（但 %% 是合法的）
            idx = line.find('%')
            if idx >= 0:
                # 检查 %% 类型
                count = 0
                for j in range(idx):
                    if line[j] == '%':
                        count += 1
                if count % 2 == 0:
                    line = line[:idx]
            result.append(line)
        return '\n'.join(result)

    def _extract_braced_content(self, text: str) -> str:
        """提取大括号内容"""
        if not text or text[0] != '{':
            return ""

        depth = 0
        start = 0
        content = []

        for i, char in enumerate(text):
            if char == '{':
                if depth == 0:
                    start = i + 1
                depth += 1
            elif char == '}':
                depth -= 1
                if depth == 0:
                    content.append(text[start:i])
                    break

        return ''.join(content) if content else ""

    def _extract_environment(self, content: str, start: int, env_name: str) -> str:
        """提取环境内容"""
        begin_pattern = f'\\begin{{{env_name}}}'
        end_pattern = f'\\end{{{env_name}}}'

        begin_idx = content.find(begin_pattern, start)
        if begin_idx < 0:
            return ""

        end_idx = content.find(end_pattern, begin_idx)
        if end_idx < 0:
            return ""

        return content[begin_idx + len(begin_pattern):end_idx]

    def _parse_table(self, content: str) -> Optional[TableElement]:
        """解析表格"""
        rows = []
        lines = content.split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 移除 \hline
            line = re.sub(r'\\hline', '', line).strip()
            if not line:
                continue

            # 分割单元格 (& 分隔)
            cells = [c.strip() for c in line.split('&')]
            # 清理单元格内容
            cells = [self._clean_cell_text(c) for c in cells]
            rows.append(cells)

        if rows:
            return TableElement(rows=rows, has_header=True)
        return None

    def _clean_cell_text(self, text: str) -> str:
        """清理单元格文本"""
        # 移除 $...$ 数学公式标记（转为普通文本）
        text = re.sub(r'\$([^$]*)\$', r'\1', text)
        # 清理多余空格
        text = ' '.join(text.split())
        return text

    def _parse_list(self, content: str, list_type: str) -> List[TextElement]:
        """解析列表项"""
        items = []

        # 分割列表项
        pattern = r'\\item(?:\s*\[([^\]]*)\])?\s*'
        parts = re.split(pattern, content)

        for part in parts:
            if part is None:
                continue
            part = part.strip()
            if not part or part.startswith('['):
                continue

            # 清理文本
            part = self._clean_cell_text(part)
            if part:
                items.append(TextElement(content=part))

        return items


def parse_latex(tex: str) -> LaTeXDocument:
    """便捷函数：解析 LaTeX 文本"""
    parser = LaTeXParser()
    return parser.parse(tex)