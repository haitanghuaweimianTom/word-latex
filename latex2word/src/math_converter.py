"""
LaTeX 数学公式到 OMML (Office Math Markup Language) 转换器
"""

from __future__ import annotations

import re
from typing import Optional


class LaTeXToOMMLConverter:
    """
    将 LaTeX 数学公式转换为 OMML (Word 内置格式)

    支持:
    - 基本运算: +, -, *, /, =
    - 上下标: ^{}, _{}
    - 分式: \frac{}{}
    - 根号: \sqrt{}, \sqrt[]{}
    - 希腊字母: \alpha, \beta, \gamma, etc.
    - 大型运算符: \sum, \prod, \int
    """

    # 希腊字母映射
    GREEK_LETTERS = {
        'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ',
        'epsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ',
        'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ',
        'nu': 'ν', 'xi': 'ξ', 'pi': 'π', 'rho': 'ρ',
        'sigma': 'σ', 'tau': 'τ', 'upsilon': 'υ', 'phi': 'φ',
        'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
        'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ',
        'Xi': 'Ξ', 'Pi': 'Π', 'Sigma': 'Σ', 'Upsilon': 'Υ',
        'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω',
    }

    # 特殊函数
    SPECIAL_FUNCTIONS = {
        'sin': 'sin', 'cos': 'cos', 'tan': 'tan',
        'log': 'log', 'ln': 'ln', 'exp': 'exp',
        'lim': 'lim', 'max': 'max', 'min': 'min',
    }

    def convert(self, latex: str) -> str:
        """
        将 LaTeX 公式转换为可读文本

        对于复杂的数学公式，我们返回一个简化的文本表示。
        Word 的 OMML 支持有限，我们采用简化方法。

        Args:
            latex: LaTeX 数学公式字符串

        Returns:
            简化后的公式文本
        """
        if not latex:
            return ""

        # 预处理
        latex = latex.strip()

        # 处理分式 \frac{}{}
        latex = self._process_fractions(latex)

        # 处理根号 \sqrt{} 和 \sqrt[]{}
        latex = self._process_sqrt(latex)

        # 处理上下标 ^{} 和 _{}
        latex = self._process_super_sub(latex)

        # 处理希腊字母
        latex = self._process_greek(latex)

        # 处理特殊函数
        latex = self._process_functions(latex)

        # 处理大括号包裹的内容
        latex = self._process_braces(latex)

        # 处理运算符
        latex = self._process_operators(latex)

        # 清理多余空格
        latex = ' '.join(latex.split())

        return latex

    def _process_fractions(self, latex: str) -> str:
        """处理分式 \frac{num}{den}"""
        pattern = r'\\frac\{([^{}]*)\}\{([^{}]*)\}'

        while True:
            match = re.search(pattern, latex)
            if not match:
                break

            num = match.group(1)
            den = match.group(2)

            # 替换为 (num)/(den) 格式
            replacement = f'({num})/({den})'
            start, end = match.span()
            latex = latex[:start] + replacement + latex[end:]

        return latex

    def _process_sqrt(self, latex: str) -> str:
        """处理根号 \sqrt{} 和 \sqrt[]{}"""
        # 处理 \sqrt[]{}
        pattern = r'\\sqrt\[([^\]]*)\]\{([^{}]*)\}'
        while True:
            match = re.search(pattern, latex)
            if not match:
                break

            index = match.group(1)
            content = match.group(2)
            replacement = f'({content})^(1/{index})'
            start, end = match.span()
            latex = latex[:start] + replacement + latex[end:]

        # 处理 \sqrt{}
        pattern = r'\\sqrt\{([^{}]*)\}'
        while True:
            match = re.search(pattern, latex)
            if not match:
                break

            content = match.group(1)
            replacement = f'sqrt({content})'
            start, end = match.span()
            latex = latex[:start] + replacement + latex[end:]

        return latex

    def _process_super_sub(self, latex: str) -> str:
        """处理上下标 ^{} 和 _{}"""
        # 处理上标
        pattern = r'\^{([^{}]*)}'
        while True:
            match = re.search(pattern, latex)
            if not match:
                break

            content = match.group(1)
            replacement = f'^({content})'
            start, end = match.span()
            latex = latex[:start] + replacement + latex[end:]

        # 处理下标
        pattern = r'_{-'
        latex = latex.replace('_{-', '_(')
        pattern = r'_\{([^{}]*)\}'
        while True:
            match = re.search(pattern, latex)
            if not match:
                break

            content = match.group(1)
            replacement = f'_({content})'
            start, end = match.span()
            latex = latex[:start] + replacement + latex[end:]

        return latex

    def _process_greek(self, latex: str) -> str:
        """处理希腊字母"""
        for greek, char in self.GREEK_LETTERS.items():
            latex = latex.replace(f'\\{greek}', char)
        return latex

    def _process_functions(self, latex: str) -> str:
        """处理特殊函数"""
        for func, name in self.SPECIAL_FUNCTIONS.items():
            latex = latex.replace(f'\\{func}', func)
        return latex

    def _process_braces(self, latex: str) -> str:
        """处理大括号包裹的内容"""
        # 将 {abc} 转换为 (abc)
        while True:
            match = re.search(r'\{([^{}]*)\}', latex)
            if not match:
                break
            content = match.group(1)
            start, end = match.span()
            latex = latex[:start] + '(' + content + ')' + latex[end:]

        return latex

    def _process_operators(self, latex: str) -> str:
        """处理运算符"""
        # 替换 LaTeX 运算符
        replacements = {
            '\\times': '×',
            '\\div': '÷',
            '\\pm': '±',
            '\\leq': '≤',
            '\\geq': '≥',
            '\\neq': '≠',
            '\\sum': '∑',
            '\\prod': '∏',
            '\\int': '∫',
            '\\infty': '∞',
            '\\cdot': '·',
            '\\ldots': '…',
        }

        for latex_op, char in replacements.items():
            latex = latex.replace(latex_op, char)

        return latex


def convert_latex_math_to_omml(latex: str) -> str:
    """便捷函数：将 LaTeX 数学公式转换为简化文本"""
    converter = LaTeXToOMMLConverter()
    return converter.convert(latex)