"""
Word → LaTeX 转换引擎
负责解析 Word 文档结构并转换为 LaTeX 代码
"""

from __future__ import annotations

import os
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from mathml_to_latex import MathMLToLaTeX
from table_converter import TableConverter


@dataclass
class ConversionOptions:
    """转换选项"""
    img_dir: str = "./figures"
    package_mode: bool = False
    verbose: bool = False

    @property
    def output_dir(self) -> str:
        return os.path.dirname(self.img_dir) or "."


@dataclass
class Word2LaTeXConverter:
    """Word 到 LaTeX 转换器"""

    options: ConversionOptions = field(default_factory=ConversionOptions)
    math_converter: MathMLToLaTeX = field(default_factory=MathMLToLaTeX)
    table_converter: TableConverter = field(default_factory=TableConverter)

    _inline_math_count: int = 0
    _display_math_count: int = 0
    _image_count: int = 0
    _extracted_images: set = field(default_factory=set)

    # Word 标题样式名 → LaTeX 命令映射
    HEADING_STYLES = {
        'Heading 1': ('\\section', False),
        'Heading 2': ('\\subsection', False),
        'Heading 3': ('\\subsubsection', False),
        'Heading 4': ('\\paragraph', False),
        'Heading 5': ('\\subparagraph', False),
        'Heading 6': ('\\subparagraph', True),
    }

    def convert(self, docx_path: str) -> str:
        """
        转换 Word 文档为 LaTeX 代码

        Args:
            docx_path: Word 文档路径

        Returns:
            LaTeX 代码字符串
        """
        doc = Document(docx_path)

        # 提取图片
        self._extract_images(docx_path, doc)

        # 转换文档内容
        parts = []

        if not self.options.package_mode:
            parts.append(self._generate_preamble())

        # 处理文档中的元素
        for element in doc.element.body:
            result = self._convert_element(element)
            if result:
                parts.append(result)

        if not self.options.package_mode:
            parts.append("\\end{document}")

        return '\n\n'.join(filter(None, parts))

    def _generate_preamble(self) -> str:
        """生成 LaTeX 文档导言区"""
        packages = [
            "\\usepackage{amsmath,amssymb}",
            "\\usepackage{graphicx}",
            "\\usepackage{booktabs}",
            "\\usepackage{geometry}",
            "\\usepackage{enumerate}",
        ]

        return """\\documentclass{article}
\\usepackage[utf8]{inputenc}
\\usepackage[T1]{fontenc}

""" + '\n'.join(packages) + """

\\begin{document}
"""

    def _convert_element(self, element) -> Optional[str]:
        """转换单个文档元素"""
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag

        if tag == 'p':
            return self._convert_paragraph(element)
        elif tag == 'tbl':
            return self._convert_table(element)

        return None

    def _convert_paragraph(self, p_element) -> Optional[str]:
        """转换段落元素"""
        # 获取段落样式
        style_name = self._get_paragraph_style(p_element)

        # 获取文本内容
        text_content = []
        is_display_math = False

        for child in p_element.iter():
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'r':  # Run (文字片段)
                text = self._convert_run(child)
                if text:
                    text_content.append(text)

            elif tag == 'm':  # Math (Office Math)
                math_result = self._convert_math(child)
                if math_result:
                    text_content.append(math_result)
                    if math_result.startswith('\\['):
                        is_display_math = True

        if not text_content:
            return None

        full_text = ''.join(text_content).strip()
        if not full_text:
            return None

        # 根据样式确定 LaTeX 命令
        if style_name in self.HEADING_STYLES:
            latex_cmd, need_subparagraph = self.HEADING_STYLES[style_name]
            # 清理文本中的特殊字符
            clean_text = self._clean_latex_text(full_text)

            if need_subparagraph:
                cmd = f"\\{latex_cmd}"
                return f"{cmd}{{\\subsubsection*{{{clean_text}}}}}"

            # 添加星号避免编号，或使用标准命令
            if style_name in ('Heading 1', 'Heading 2'):
                return f"\\{latex_cmd}*{{{clean_text}}}"
            return f"\\{latex_cmd}{{{clean_text}}}"

        # 处理列表项
        if self._is_list_item(p_element):
            bullet = self._get_list_bullet(p_element)
            return f"\\item {full_text}"

        # 普通段落
        if is_display_math:
            return full_text  # 已经是 display math 格式

        # 检测是否应以行间公式结束（用户可能希望换行）
        return f"{full_text}"

    def _convert_run(self, r_element) -> str:
        """转换文字片段 (Run)"""
        texts = []

        for child in r_element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 't':  # Text
                texts.append(child.text or '')
            elif tag == 'br':  # Line break
                texts.append('\\\\')
            elif tag == 'tab':
                texts.append('\\hspace{1em}')

        text = ''.join(texts)

        # 检查格式
        is_bold = self._run_has_style(r_element, 'b')
        is_italic = self._run_has_style(r_element, 'i')

        # 处理数学公式中的文本（保持原样）
        if self.math_converter.is_in_math_context:
            return text

        # 应用格式
        if is_bold and is_italic:
            text = f"\\textbf{{\\textit{{{text}}}}}"
        elif is_bold:
            text = f"\\textbf{{{text}}}"
        elif is_italic:
            text = f"\\textit{{{text}}}"

        return text

    def _convert_math(self, m_element) -> str:
        """转换数学公式 (MathML → LaTeX)"""
        # 提取 MathML
        math_xml = self._extract_mathml(m_element)

        if math_xml:
            return self.math_converter.convert(math_xml)

        return ''

    def _extract_mathml(self, m_element) -> Optional[str]:
        """从 Word Math 元素中提取 MathML"""
        try:
            # 遍历子元素找 MathBase
            for child in m_element.iter():
                if 'MathBase' in child.tag or 'oMath' in child.tag:
                    # 尝试构建 MathML
                    mathml = self._build_mathml_from_element(child)
                    if mathml:
                        return mathml
        except Exception:
            pass

        return None

    def _build_mathml_from_element(self, element) -> Optional[str]:
        """从元素构建 MathML 字符串"""
        try:
            import xml.etree.ElementTree as ET
            xml_str = ET.tostring(element, encoding='unicode')

            # 简化处理：如果包含 MathML 命名空间标记
            if 'xmlns="http://www.w3.org/1998/Math/MathML"' in xml_str or \
               'xmlns:m' in xml_str:
                return xml_str

            # 检查是否有 m:oMath
            if 'oMath' in element.tag or 'm:' in str(element.tags if hasattr(element, 'tags') else ''):
                # 构建基本 MathML 结构
                return self._create_mathml_from_omml(element)
        except Exception:
            pass

        return None

    def _create_mathml_from_omml(self, element) -> str:
        """
        从 Office Math ML (OMML) 转换为 MathML
        这是关键转换逻辑
        """
        from xml.etree import ElementTree as ET

        # 创建 MathML 根元素
        mathml = ET.Element('math', xmlns='http://www.w3.org/1998/Math/MathML')

        # 递归转换 OMML 元素
        self._convert_omml_node(element, mathml)

        return ET.tostring(mathml, encoding='unicode')

    def _convert_omml_node(self, omml_elem, mathml_parent):
        """递归转换 OMML 到 MathML 节点"""
        from xml.etree import ElementTree as ET

        tag = omml_elem.tag.split('}')[-1] if '}' in omml_elem.tag else omml_elem.tag

        # OMML 到 MathML 的映射规则
        omml_to_mathml = {
            'oMath': ('mrow', {}),
            'oSup': ('msup', {}),
            'oSub': ('msub', {}),
            'oSupSub': ('msubsup', {}),
            'oSupPr': ('msup', {}),
            'oSubPr': ('msub', {}),
            'oDen': ('mfrac', {}),
            'oNum': ('mfrac', {}),
            'acc': ('mover', {}),
            'accPr': ('mover', {}),
            'bar': ('mover', {}),
            'barPr': ('mover', {}),
            'd': ('mrow', {}),  # d = d (trigonometric degree)
            'func': ('mi', {}),
            'funcPr': ('mi', {}),
            'e': ('mi', {}),
            'r': ('mrow', {}),  # r = argument/run
            't': ('mn', {}),    # t = text/number
            'str': ('mn', {}),
        }

        if tag in omml_to_mathml:
            mathml_tag, attrs = omml_to_mathml[tag]
            new_elem = ET.SubElement(mathml_parent, mathml_tag, attrs)

            # 处理文本内容
            if tag == 't' or tag == 'str':
                if omml_elem.text:
                    new_elem.text = omml_elem.text

            # 递归处理子元素
            for child in omml_elem:
                self._convert_omml_node(child, new_elem)

    def _get_paragraph_style(self, p_element) -> str:
        """获取段落样式名称"""
        try:
            pPr = p_element.find(qn('w:pPr'))
            if pPr is not None:
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None:
                    return pStyle.get(qn('w:val'), '')
        except Exception:
            pass
        return ''

    def _run_has_style(self, r_element, style_char: str) -> bool:
        """检查 Run 是否具有特定样式（加粗/斜体）"""
        try:
            rPr = r_element.find(qn('w:rPr'))
            if rPr is not None:
                style_elem = rPr.find(qn(f'w:{style_char}'))
                return style_elem is not None
        except Exception:
            pass
        return False

    def _is_list_item(self, p_element) -> bool:
        """检查是否是列表项"""
        try:
            pPr = p_element.find(qn('w:pPr'))
            if pPr is not None:
                numPr = pPr.find(qn('w:numPr'))
                return numPr is not None
        except Exception:
            pass
        return False

    def _get_list_bullet(self, p_element) -> str:
        """获取列表项目符号"""
        # 默认返回 bullet
        return "\\item"

    def _clean_latex_text(self, text: str) -> str:
        """清理文本中的特殊 LaTeX 字符"""
        # 转义特殊字符
        special_chars = {
            '#': '\\#',
            '$': '\\$',
            '%': '\\%',
            '&': '\\&',
            '{': '\\{',
            '}': '\\}',
            '_': '\\_',
            '^': '\\^{}',
            '~': '\\~{}',
            '\\': '\\textbackslash{}',
        }

        for char, escaped in special_chars.items():
            text = text.replace(char, escaped)

        return text

    def _convert_table(self, tbl_element) -> str:
        """转换表格元素"""
        return self.table_converter.convert(tbl_element)

    def _extract_images(self, docx_path: str, doc: Document):
        """从 docx 中提取图片"""
        if not os.path.exists(self.options.img_dir):
            os.makedirs(self.options.img_dir, exist_ok=True)

        # 使用 zipfile 直接读取 docx
        with zipfile.ZipFile(docx_path, 'r') as zf:
            # 列出所有文件
            for name in zf.namelist():
                if name.startswith('word/media/'):
                    # 提取图片
                    base_name = os.path.basename(name)
                    dest_path = os.path.join(self.options.img_dir, base_name)

                    if base_name not in self._extracted_images:
                        with zf.open(name) as src, open(dest_path, 'wb') as dst:
                            dst.write(src.read())
                        self._extracted_images.add(base_name)

        # 记录文档中的图片引用
        for rel in doc.part.rels.values():
            if "image" in rel.reltype:
                pass  # 图片已在 zip 中提取

    def get_image_ref(self, image_name: str) -> str:
        """获取图片的 LaTeX 引用"""
        return f"\\includegraphics[width=0.8\\textwidth]{{{self.options.img_dir}/{image_name}}}"