"""
LaTeX 表格到 Word 表格转换器
"""

from __future__ import annotations

from typing import List, Tuple


class LaTeXTableToWordConverter:
    """
    将 LaTeX tabular 环境转换为 Word 表格

    支持:
    - 基本表格
    - 多列表格
    - 带分隔线的表格
    """

    def convert(self, rows: List[List[str]]) -> List[List[str]]:
        """
        转换表格数据

        Args:
            rows: LaTeX 表格行数据

        Returns:
            清理后的表格数据
        """
        cleaned_rows = []

        for row in rows:
            cleaned_row = []
            for cell in row:
                # 清理单元格内容
                cleaned_cell = self._clean_cell(cell)
                cleaned_row.append(cleaned_cell)
            cleaned_rows.append(cleaned_row)

        return cleaned_rows

    def _clean_cell(self, text: str) -> str:
        """
        清理单元格内容

        - 移除 $...$ 数学公式标记
        - 清理多余空白
        - 处理转义字符
        """
        if not text:
            return ""

        # 移除数学公式标记
        import re
        text = re.sub(r'\$([^$]*)\$', r'\1', text)
        text = re.sub(r'\$\$([^$]*)\$\$', r'\1', text)

        # 移除 LaTeX 命令标记
        text = re.sub(r'\\[a-zA-Z]+', '', text)

        # 移除多余空白
        text = ' '.join(text.split())

        return text


def convert_latex_table(rows: List[List[str]]) -> List[List[str]]:
    """便捷函数：转换 LaTeX 表格"""
    converter = LaTeXTableToWordConverter()
    return converter.convert(rows)