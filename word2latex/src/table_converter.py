"""
Word 表格 → LaTeX 表格环境转换器
"""

from __future__ import annotations

from typing import List
from docx.oxml.ns import qn


class TableConverter:
    """将 Word 表格转换为 LaTeX tabular 环境"""

    def convert(self, tbl_element) -> str:
        try:
            rows = self._get_rows(tbl_element)
            if not rows:
                return ""
            num_cols = len(self._get_cells(rows[0]))
            return self._generate_latex_table(rows, num_cols)
        except Exception as e:
            return "% Table conversion failed: " + str(e)

    def _get_rows(self, tbl) -> List:
        rows = []
        for child in tbl:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "tr":
                rows.append(child)
        return rows

    def _get_cells(self, row) -> List:
        cells = []
        for child in row:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "tc":
                cells.append(child)
        return cells

    def _get_cell_content(self, cell) -> str:
        texts = []
        for child in cell:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                para_text = self._get_paragraph_text(child)
                if para_text:
                    texts.append(para_text)
        return " ".join(texts)

    def _get_paragraph_text(self, p_element) -> str:
        text_parts = []
        for child in p_element.iter():
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "t":
                text_parts.append(child.text or "")
        return "".join(text_parts)

    def _generate_latex_table(self, rows: List, num_cols: int) -> str:
        bs = chr(92)
        lines = []
        lines.append(bs + "begin{tabular}{|" + "|".join(["c"] * num_cols) + "|}")
        lines.append(bs + "hline")

        for i, row in enumerate(rows):
            cells = self._get_cells(row)
            cell_contents = [self._clean_cell_content(self._get_cell_content(c)) for c in cells]
            lines.append(" & ".join(cell_contents))
            if i < len(rows) - 1:
                lines.append(bs + "hline")

        lines.append(bs + "hline")
        lines.append(bs + "end{tabular}")
        return "\n".join(lines)

    def _clean_cell_content(self, content: str) -> str:
        """清理单元格内容中的 LaTeX 特殊字符"""
        PLACEHOLDER = chr(0x0001F600)
        bs = chr(92)
        
        content = content.replace(bs, PLACEHOLDER)
        content = content.replace("{", bs + "{")
        content = content.replace("}", bs + "}")
        content = content.replace("&", bs + "&")
        content = content.replace("%", bs + "%")
        content = content.replace("$", bs + "$")
        content = content.replace("#", bs + "#")
        content = content.replace("_", bs + "_")
        content = content.replace("~", bs + "textasciitilde{}")
        content = content.replace("^", bs + "textasciicircum{}")
        content = content.replace(PLACEHOLDER, bs + "textbackslash{}")
        content = content.replace(chr(10), " " + bs + bs + " ")
        
        return content

    def convert_with_title(self, tbl_element, title: str = "",
                          caption: str = "", label: str = "") -> str:
        bs = chr(92)
        parts = []
        if caption:
            parts.append(bs + "caption{" + caption + "}")
        if label:
            parts.append(bs + "label{" + label + "}")
        parts.append(self.convert(tbl_element))
        if title:
            parts.append("% Table: " + title)
        return "\n".join(parts)
