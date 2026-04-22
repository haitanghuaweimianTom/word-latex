"""
MathML / OMML → LaTeX 转换器
将 Word 文档中的数学公式转换为 LaTeX 格式
"""

from __future__ import annotations

import re
from typing import Optional, Dict, Tuple


class MathMLToLaTeX:
    """
    MathML 或 Office Math ML (OMML) 到 LaTeX 的转换器

    支持的数学元素:
    - 基本运算符: +, -, *, /, =, <, >
    - 上下标: superscript, subscript
    - 分式: fraction
    - 根号: sqrt, nth-root
    - 矩阵: matrix, determinants
    - 希腊字母: alpha, beta, gamma, etc.
    - 大型运算符: sum, product, integral
    - 箭头和关系符
    """

    # 希腊字母映射 (小写)
    GREEK_LETTERS = {
        'alpha': r'\alpha',
        'beta': r'\beta',
        'gamma': r'\gamma',
        'delta': r'\delta',
        'epsilon': r'\epsilon',
        'varepsilon': r'\varepsilon',
        'zeta': r'\zeta',
        'eta': r'\eta',
        'theta': r'\theta',
        'iota': r'\iota',
        'kappa': r'\kappa',
        'lambda': r'\lambda',
        'mu': r'\mu',
        'nu': r'\nu',
        'xi': r'\xi',
        'pi': r'\pi',
        'rho': r'\rho',
        'sigma': r'\sigma',
        'tau': r'\tau',
        'upsilon': r'\upsilon',
        'phi': r'\phi',
        'chi': r'\chi',
        'psi': r'\psi',
        'omega': r'\omega',
        # 大写希腊字母
        'Gamma': r'\Gamma',
        'Delta': r'\Delta',
        'Theta': r'\Theta',
        'Lambda': r'\Lambda',
        'Xi': r'\Xi',
        'Pi': r'\Pi',
        'Sigma': r'\Sigma',
        'Upsilon': r'\Upsilon',
        'Phi': r'\Phi',
        'Psi': r'\Psi',
        'Omega': r'\Omega',
    }

    # 特殊函数名
    SPECIAL_FUNCTIONS = {
        'sin': r'\sin',
        'cos': r'\cos',
        'tan': r'\tan',
        'cot': r'\cot',
        'sec': r'\sec',
        'csc': r'\csc',
        'log': r'\log',
        'ln': r'\ln',
        'exp': r'\exp',
        'lim': r'\lim',
        'max': r'\max',
        'min': r'\min',
        'sup': r'\sup',
        'inf': r'\inf',
        'det': r'\det',
        'det': r'\det',
        'dim': r'\dim',
        'ker': r'\ker',
        'deg': r'\deg',
    }

    # 运算符映射
    OPERATORS = {
        '+': '+',
        '-': '-',
        '·': r'\cdot',
        '×': r'\times',
        '÷': r'\div',
        '=': '=',
        '≠': r'\neq',
        '≈': r'\approx',
        '≡': r'\equiv',
        '<': '<',
        '>': '>',
        '≤': r'\leq',
        '≥': r'\geq',
        '±': r'\pm',
        '∓': r'\mp',
        '∝': r'\propto',
    }

    def __init__(self):
        self.is_in_math_context = False
        self._buffer: list[str] = []

    def convert(self, xml_content: str) -> str:
        """
        将 MathML/OMML 转换为 LaTeX

        Args:
            xml_content: MathML 或 OMML XML 字符串

        Returns:
            LaTeX 公式字符串
        """
        if not xml_content or len(xml_content.strip()) == 0:
            return ''

        # 检测是否为行间公式（通常以 \[ 或 display math 结尾）
        is_display = self._detect_display_math(xml_content)

        try:
            result = self._parse_mathml(xml_content)
        except Exception as e:
            # 回退到基于文本的转换
            result = self._fallback_conversion(xml_content)

        if is_display:
            return f"\\[{result}\\]"
        else:
            return f"\\({result}\\)"

    def _detect_display_math(self, xml_content: str) -> bool:
        """检测是否为行间公式"""
        # 行间公式通常在块级段落中
        return False  # 默认按行内处理，除非明确检测到

    def _parse_mathml(self, xml_content: str) -> str:
        """解析 MathML 并转换为 LaTeX"""
        import xml.etree.ElementTree as ET

        # 解析 XML
        try:
            root = ET.fromstring(xml_content)
        except ET.ParseError:
            return self._fallback_conversion(xml_content)

        # 检测根元素类型
        if 'MathML' in xml_content or 'math' in root.tag:
            return self._convert_mathml_node(root)
        elif 'omath' in xml_content.lower() or 'oMath' in xml_content:
            return self._convert_omml_node(root)

        return str(xml_content)

    def _convert_mathml_node(self, element) -> str:
        """递归转换 MathML 节点"""
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag

        converters = {
            'math': self._convert_math,
            'mrow': self._convert_row,
            'mfrac': self._convert_fraction,
            'msqrt': self._convert_sqrt,
            'mroot': self._convert_nthroot,
            'msup': self._convert_superscript,
            'msub': self._convert_subscript,
            'msubsup': self._convert_subsuperscript,
            'mi': self._convert_identifier,
            'mn': self._convert_number,
            'mo': self._convert_operator,
            'mtext': self._convert_text,
            'mspace': self._convert_space,
            'mrow': self._convert_row,
            'mtable': self._convert_table,
            'mtr': self._convert_row,
            'mtd': self._convert_row,
            'mover': self._convert_over,
            'munder': self._convert_under,
            'munderover': self._convert_underover,
            'menclose': self._convert_enclose,
        }

        converter = converters.get(tag)
        if converter:
            return converter(element)

        # 未知的元素，返回文本内容
        return ''.join(self._convert_text_node(child) for child in element)

    def _convert_text_node(self, element) -> str:
        """转换文本节点"""
        if element.text:
            return element.text
        return ''

    def _convert_math(self, element) -> str:
        """转换 math 根元素"""
        children = list(element)
        if children:
            return self._convert_mathml_node(children[0])
        return ''

    def _convert_row(self, element) -> str:
        """转换 row (mrow) - 水平排列"""
        parts = []
        for child in element:
            result = self._convert_mathml_node(child)
            if result:
                parts.append(result)

        return ' '.join(parts)

    def _convert_fraction(self, element) -> str:
        """转换分式 (mfrac)"""
        num = ''
        den = ''

        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'mrow':
                # 根据顺序判断是分子还是分母
                # 在标准 MathML 中，第一个是分子，第二个是分母
                if not num:
                    num = self._convert_row(child)
                else:
                    den = self._convert_row(child)

        if num and den:
            return f"\\frac{{{num}}}{{{den}}}"
        return num or den or ''

    def _convert_sqrt(self, element) -> str:
        """转换平方根 (msqrt)"""
        content = self._get_single_child_content(element)
        if content:
            return f"\\sqrt{{{content}}}"
        return ''

    def _convert_nthroot(self, element) -> str:
        """转换 n 次方根 (mroot)"""
        children = list(element)
        if len(children) >= 2:
            base = self._convert_mathml_node(children[0])
            index = self._convert_mathml_node(children[1])
            return f"\\sqrt[{index}]{{{base}}}"
        return ''

    def _convert_superscript(self, element) -> str:
        """转换上标 (msup)"""
        children = list(element)
        if len(children) >= 2:
            base = self._convert_mathml_node(children[0])
            sup = self._convert_mathml_node(children[1])
            return f"{{{base}}}^{{{sup}}}"
        return ''

    def _convert_subscript(self, element) -> str:
        """转换下标 (msub)"""
        children = list(element)
        if len(children) >= 2:
            base = self._convert_mathml_node(children[0])
            sub = self._convert_mathml_node(children[1])
            return f"{{{base}}}_{{{sub}}}"
        return ''

    def _convert_subsuperscript(self, element) -> str:
        """转换上下标 (msubsup)"""
        children = list(element)
        if len(children) >= 3:
            base = self._convert_mathml_node(children[0])
            sub = self._convert_mathml_node(children[1])
            sup = self._convert_mathml_node(children[2])
            return f"{{{base}}}_{{{sub}}}^{{{sup}}}"
        elif len(children) == 2:
            # 可能是上标或下标
            return self._convert_superscript(element)
        return ''

    def _convert_identifier(self, element) -> str:
        """转换标识符 (mi) - 变量名、函数名"""
        text = element.text or ''

        # 检查是否是希腊字母
        if text.lower() in self.GREEK_LETTERS:
            return self.GREEK_LETTERS.get(text, text)

        # 检查是否是特殊函数
        if text.lower() in self.SPECIAL_FUNCTIONS:
            return self.SPECIAL_FUNCTIONS[text.lower()]

        # 普通标识符用斜体
        if len(text) == 1:
            return f"{{{text}}}"  # 单字符斜体

        # 多字符可能是函数名，保持 upright
        return f"\\text{{{text}}}"

    def _convert_number(self, element) -> str:
        """转换数字 (mn)"""
        return element.text or ''

    def _convert_operator(self, element) -> str:
        """转换运算符 (mo)"""
        text = element.text or ''

        # 检查预定义运算符
        if text in self.OPERATORS:
            return self.OPERATORS[text]

        # 大型运算符
        large_operators = {
            '∑': r'\sum',
            '∏': r'\prod',
            '∫': r'\int',
            '∬': r'\iint',
            '∮': r'\oint',
            '∂': r'\partial',
            '∇': r'\nabla',
        }

        if text in large_operators:
            return large_operators[text]

        # 返回原文本
        return text

    def _convert_text(self, element) -> str:
        """转换文本 (mtext)"""
        return element.text or ''

    def _convert_space(self, element) -> str:
        """转换空格 (mspace)"""
        return ' '

    def _convert_table(self, element) -> str:
        """转换表格 (mtable) - 矩阵等"""
        rows = []
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'mtr':
                cells = []
                for cell in child:
                    cell_tag = cell.tag.split('}')[-1] if '}' in cell.tag else cell.tag
                    if cell_tag == 'mtd':
                        content = self._get_single_child_content(cell)
                        cells.append(content)
                if cells:
                    rows.append(' & '.join(cells))

        if rows:
            row_str = ' \\\\ '.join(rows)
            return f"\\begin{{matrix}}{row_str}\\end{{matrix}}"

        return ''

    def _convert_over(self, element) -> str:
        """转换上标装饰 (mover) - 如向量、波浪号"""
        children = list(element)
        if len(children) >= 2:
            base = self._get_single_child_content(children[0])
            accent = self._get_single_child_content(children[1])

            # 判断装饰类型
            accent_map = {
                '→': r'\vec',
                '←': r'\vec',
                '↔': r'\vec',
                '~': r'\tilde',
                '^': r'\hat',
                '¯': r'\bar',
            }

            if accent in accent_map:
                return f"{accent_map[accent]}{{{base}}}"
            elif accent == '→':
                return f"\\overrightarrow{{{base}}}"
            elif accent == '←':
                return f"\\overleftarrow{{{base}}}"
            elif accent == '↔':
                return f"\\overleftrightarrow{{{base}}}"

        return ''

    def _convert_under(self, element) -> str:
        """转换下标装饰 (munder)"""
        children = list(element)
        if len(children) >= 2:
            base = self._get_single_child_content(children[0])
            under = self._get_single_child_content(children[1])
            return f"\\underset{{{under}}}{{{base}}}"

        return ''

    def _convert_underover(self, element) -> str:
        """转换上下标装饰 (munderover)"""
        children = list(element)
        if len(children) >= 3:
            base = self._get_single_child_content(children[0])
            under = self._get_single_child_content(children[1])
            over = self._get_single_child_content(children[2])
            return "\\underset{" + under + "}{\\overset{" + over + "}{" + base + "}}"

        return ''

    def _convert_enclose(self, element) -> str:
        """转换环绕装饰 (menclose)"""
        notation = element.get('notation', '')

        content = self._get_single_child_content(element)

        if 'box' in notation or 'border' in notation:
            return f"\\boxed{{{content}}}"
        elif 'radical' in notation:
            return f"\\sqrt{{{content}}}"

        return content

    def _get_single_child_content(self, element) -> str:
        """获取单个子元素的内容"""
        children = list(element)
        if children:
            return self._convert_mathml_node(children[0])
        return element.text or ''

    def _convert_omml_node(self, element) -> str:
        """转换 Office Math ML (OMML) 节点"""
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag

        omml_converters = {
            'oMath': self._convert_omml_math,
            'oSup': self._convert_omml_sup,
            'oSub': self._convert_omml_sub,
            'oSupSub': self._convert_omml_subsup,
            'd': self._convert_omml_degree,
            'func': self._convert_omml_function,
            'e': self._convert_omml_element,
            'r': self._convert_omml_run,
            't': self._convert_omml_text,
        }

        converter = omml_converters.get(tag)
        if converter:
            return converter(element)

        # 处理子元素
        parts = []
        for child in element:
            result = self._convert_omml_node(child)
            if result:
                parts.append(result)

        return ''.join(parts)

    def _convert_omml_math(self, element) -> str:
        """转换 oMath 元素"""
        parts = []
        for child in element:
            result = self._convert_omml_node(child)
            if result:
                parts.append(result)
        return ' '.join(parts)

    def _convert_omml_sup(self, element) -> str:
        """转换上标"""
        children = list(element)
        if len(children) >= 2:
            base = self._get_omml_text(children[0])
            sup = self._get_omml_text(children[1])
            return f"{{{base}}}^{{{sup}}}"
        return ''

    def _convert_omml_sub(self, element) -> str:
        """转换下标"""
        children = list(element)
        if len(children) >= 2:
            base = self._get_omml_text(children[0])
            sub = self._get_omml_text(children[1])
            return f"{{{base}}}_{{{sub}}}"
        return ''

    def _convert_omml_subsup(self, element) -> str:
        """转换上下标"""
        children = list(element)
        base = ''
        sub = ''
        sup = ''

        for child in children:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            text = self._get_omml_text(child)

            if tag == 'e':
                if not base:
                    base = text
            elif tag == 'sub':
                sub = text
            elif tag == 'sup':
                sup = text

        if base:
            if sub and sup:
                return f"{{{base}}}_{{{sub}}}^{{{sup}}}"
            elif sub:
                return f"{{{base}}}_{{{sub}}}"
            elif sup:
                return f"{{{base}}}^{{{sup}}}"

        return ''

    def _convert_omml_degree(self, element) -> str:
        """转换度数标记"""
        # d 标签表示 degree，如 sin 30°
        children = list(element)
        if children:
            return self._get_omml_text(children[0])
        return ''

    def _convert_omml_function(self, element) -> str:
        """转换函数"""
        # func 包含 funcPr（函数名）和 e（参数）
        children = list(element)
        func_name = ''
        func_arg = ''

        for child in children:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'fName':
                func_name = self._get_omml_text(child)
            elif tag == 'e':
                func_arg = self._convert_omml_node(child)

        if func_name:
            func_latex = self.SPECIAL_FUNCTIONS.get(func_name.lower(), func_name)

            # 带上下标的函数（如 lim, sup）
            # 检查是否有 limPr 等
            return f"\\{func_latex.lstrip('\\')}{{{func_arg}}}"

        return func_arg

    def _convert_omml_element(self, element) -> str:
        """转换基本元素 e"""
        parts = []
        for child in element:
            result = self._convert_omml_node(child)
            if result:
                parts.append(result)
        return ''.join(parts)

    def _convert_omml_run(self, element) -> str:
        """转换 run r"""
        parts = []
        for child in element:
            result = self._convert_omml_node(child)
            if result:
                parts.append(result)
        return ''.join(parts)

    def _convert_omml_text(self, element) -> str:
        """转换文本 t"""
        return element.text or ''

    def _get_omml_text(self, element) -> str:
        """获取 OMML 元素的文本内容"""
        text_parts = []
        for child in element:
            text_parts.append(self._convert_omml_node(child))
        return ''.join(text_parts) if text_parts else (element.text or '')

    def _fallback_conversion(self, text: str) -> str:
        """
        回退转换：当 XML 解析失败时使用的基于规则的转换
        处理常见的公式文本模式
        """
        import re

        # 清理文本
        text = text.strip()

        # 移除 XML 标签但保留文本
        text = re.sub(r'<[^>]+>', '', text)
        text = re.sub(r'\s+', ' ', text)

        # 替换特殊字符
        replacements = {
            '×': ' \\times ',
            '÷': ' \\div ',
            '·': ' \\cdot ',
            '√': r'\sqrt',
            '∑': r'\sum',
            '∏': r'\prod',
            '∫': r'\int',
        }

        for char, latex in replacements.items():
            text = text.replace(char, latex)

        return text

    def convert_simple(self, formula_text: str) -> str:
        """
        转换简单公式文本（用于测试）
        支持基本格式如: x^2, a_b, a/b, sqrt(x)
        """
        result = formula_text

        # 分式: a/b → \frac{a}{b}
        result = re.sub(r'(\w+)/(\w+)', r'\\frac{\1}{\2}', result)

        # 上标: x^2 → x^{2}
        result = re.sub(r'(\w)\^(\d+)', r'\1^{\2}', result)
        result = re.sub(r'(\w)\^{(\w+)}', r'\1^{\2}', result)

        # 下标: x_1 → x_{1}
        result = re.sub(r'(\w)_(\d+)', r'\1_{\2}', result)
        result = re.sub(r'(\w)_{(\w+)}', r'\1_{\2}', result)

        return result


# 全局实例
_default_converter = MathMLToLaTeX()


def convert_mathml_to_latex(xml_content: str) -> str:
    """便捷函数：将 MathML/OMML 转换为 LaTeX"""
    return _default_converter.convert(xml_content)