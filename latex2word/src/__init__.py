"""
LaTeX → Word 命令行工具
将 LaTeX 文件转换为 Word 文档
"""

__version__ = '1.0.0'

from parser import LaTeXParser, LaTeXDocument, parse_latex
from converter import LaTeXToWordConverter, ConversionOptions, convert_latex_to_word
from math_converter import LaTeXToOMMLConverter, convert_latex_math_to_omml
from table_converter import LaTeXTableToWordConverter, convert_latex_table

__all__ = [
    'LaTeXParser',
    'LaTeXDocument',
    'parse_latex',
    'LaTeXToWordConverter',
    'ConversionOptions',
    'convert_latex_to_word',
    'LaTeXToOMMLConverter',
    'convert_latex_math_to_omml',
    'LaTeXTableToWordConverter',
    'convert_latex_table',
]