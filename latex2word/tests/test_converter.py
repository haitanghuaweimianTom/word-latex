"""
latex2word 单元测试
测试 LaTeX 解析和转换功能
"""

import os
import sys
import unittest
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from parser import LaTeXParser, parse_latex, LaTeXDocument
from math_converter import LaTeXToOMMLConverter
from table_converter import LaTeXTableToWordConverter


class TestLaTeXParser(unittest.TestCase):
    """LaTeX 解析器测试"""

    def setUp(self):
        self.parser = LaTeXParser()

    def test_parse_simple_text(self):
        """测试简单文本解析"""
        latex = "Hello World"
        doc = self.parser.parse(latex)
        self.assertIsInstance(doc, LaTeXDocument)
        self.assertEqual(len(doc.elements), 1)

    def test_parse_section(self):
        """测试标题解析"""
        latex = r"\section{测试标题}"
        doc = self.parser.parse(latex)
        self.assertEqual(len(doc.elements), 1)
        self.assertEqual(doc.elements[0].name, 'section')
        self.assertEqual(doc.elements[0].args[0], '测试标题')

    def test_parse_subsection(self):
        """测试子标题解析"""
        latex = r"\subsection{子标题}"
        doc = self.parser.parse(latex)
        self.assertEqual(len(doc.elements), 1)
        self.assertEqual(doc.elements[0].name, 'subsection')

    def test_parse_bold(self):
        """测试加粗解析"""
        latex = r"\textbf{加粗文本}"
        doc = self.parser.parse(latex)
        self.assertEqual(len(doc.elements), 1)
        self.assertTrue(doc.elements[0].bold)
        self.assertEqual(doc.elements[0].content, '加粗文本')

    def test_parse_italic(self):
        """测试斜体解析"""
        latex = r"\textit{斜体文本}"
        doc = self.parser.parse(latex)
        self.assertEqual(len(doc.elements), 1)
        self.assertTrue(doc.elements[0].italic)

    def test_parse_inline_math(self):
        """测试行内公式解析"""
        latex = r"$x^2$"
        doc = self.parser.parse(latex)
        self.assertEqual(len(doc.elements), 1)
        self.assertFalse(doc.elements[0].display)

    def test_parse_display_math(self):
        """测试行间公式解析"""
        latex = r"\[x^2 + y^2\]"
        doc = self.parser.parse(latex)
        self.assertEqual(len(doc.elements), 1)
        self.assertTrue(doc.elements[0].display)

    def test_parse_itemize(self):
        """测试无序列表解析"""
        latex = r"\begin{itemize}" + "\n" + r"\item 项目一" + "\n" + r"\end{itemize}"
        doc = self.parser.parse(latex)
        self.assertGreaterEqual(len(doc.elements), 1)
        # 检查是否有任何元素包含 itemize 相关信息
        found = False
        for elem in doc.elements:
            if hasattr(elem, 'name') and elem.name == 'itemize':
                found = True
                break
        self.assertTrue(found, "Should find itemize environment")

    def test_parse_enumerate(self):
        """测试有序列表解析"""
        latex = r"\begin{enumerate}" + "\n" + r"\item 项目一" + "\n" + r"\end{enumerate}"
        doc = self.parser.parse(latex)
        self.assertGreaterEqual(len(doc.elements), 1)
        # 检查是否有任何元素包含 enumerate 相关信息
        found = False
        for elem in doc.elements:
            if hasattr(elem, 'name') and elem.name == 'enumerate':
                found = True
                break
        self.assertTrue(found, "Should find enumerate environment")


class TestMathConverter(unittest.TestCase):
    """数学公式转换器测试"""

    def setUp(self):
        self.converter = LaTeXToOMMLConverter()

    def test_simple_math(self):
        """测试简单数学公式"""
        latex = "x^2"
        result = self.converter.convert(latex)
        # 现在返回简化文本格式
        self.assertIn('x', result)
        self.assertIn('2', result)

    def test_fraction(self):
        """测试分式"""
        latex = r"\frac{a}{b}"
        result = self.converter.convert(latex)
        # 现在返回 (a)/(b) 格式
        self.assertIn('a', result)
        self.assertIn('b', result)
        self.assertIn('/', result)

    def test_greek_letters(self):
        """测试希腊字母"""
        latex = r"\alpha + \beta"
        result = self.converter.convert(latex)
        self.assertIn('α', result)
        self.assertIn('β', result)

    def test_sqrt(self):
        """测试根号"""
        latex = r"\sqrt{x}"
        result = self.converter.convert(latex)
        self.assertIn('sqrt', result)
        self.assertIn('x', result)

    def test_special_functions(self):
        """测试特殊函数"""
        latex = r"\sin(x)"
        result = self.converter.convert(latex)
        self.assertIn('sin', result)
        self.assertIn('x', result)


class TestTableConverter(unittest.TestCase):
    """表格转换器测试"""

    def setUp(self):
        self.converter = LaTeXTableToWordConverter()

    def test_clean_cell(self):
        """测试单元格清理"""
        # 测试移除数学公式标记
        text = "$x^2$"
        result = self.converter._clean_cell(text)
        self.assertEqual(result, 'x^2')

    def test_clean_whitespace(self):
        """测试空白清理"""
        text = "  多个   空格  "
        result = self.converter._clean_cell(text)
        self.assertEqual(result, '多个 空格')

    def test_convert_rows(self):
        """测试行转换"""
        rows = [['A', 'B'], ['C', 'D']]
        result = self.converter.convert(rows)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0][0], 'A')


class TestFullDocument(unittest.TestCase):
    """完整文档测试"""

    def test_parse_sample_document(self):
        """测试解析示例文档"""
        sample_tex = r"""
\documentclass{article}
\begin{document}
\section{标题}
这是文本。
\subsection{子标题}
\begin{itemize}
\item 项目
\end{itemize}
\end{document}
"""
        parser = LaTeXParser()
        doc = parser.parse(sample_tex)
        self.assertGreater(len(doc.elements), 0)


def run_tests():
    """运行所有测试"""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    suite.addTests(loader.loadTestsFromTestCase(TestLaTeXParser))
    suite.addTests(loader.loadTestsFromTestCase(TestMathConverter))
    suite.addTests(loader.loadTestsFromTestCase(TestTableConverter))
    suite.addTests(loader.loadTestsFromTestCase(TestFullDocument))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    return result.wasSuccessful()


if __name__ == '__main__':
    success = run_tests()
    sys.exit(0 if success else 1)