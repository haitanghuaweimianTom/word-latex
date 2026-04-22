"""
word2latex 单元测试
测试转换器的各项功能
"""

import os
import sys
import unittest
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from mathml_to_latex import MathMLToLaTeX
from table_converter import TableConverter


class TestMathMLToLaTeX(unittest.TestCase):
    """MathML 到 LaTeX 转换器测试"""

    def setUp(self):
        self.converter = MathMLToLaTeX()

    def test_greek_letters(self):
        """测试希腊字母转换"""
        # 小写希腊字母
        self.assertEqual(self.converter.GREEK_LETTERS['alpha'], r'\alpha')
        self.assertEqual(self.converter.GREEK_LETTERS['beta'], r'\beta')
        self.assertEqual(self.converter.GREEK_LETTERS['gamma'], r'\gamma')
        self.assertEqual(self.converter.GREEK_LETTERS['pi'], r'\pi')

        # 大写希腊字母
        self.assertEqual(self.converter.GREEK_LETTERS['Sigma'], r'\Sigma')
        self.assertEqual(self.converter.GREEK_LETTERS['Omega'], r'\Omega')

    def test_special_functions(self):
        """测试特殊函数转换"""
        self.assertEqual(self.converter.SPECIAL_FUNCTIONS['sin'], r'\sin')
        self.assertEqual(self.converter.SPECIAL_FUNCTIONS['cos'], r'\cos')
        self.assertEqual(self.converter.SPECIAL_FUNCTIONS['log'], r'\log')
        self.assertEqual(self.converter.SPECIAL_FUNCTIONS['lim'], r'\lim')

    def test_operators(self):
        """测试运算符转换"""
        self.assertEqual(self.converter.OPERATORS['×'], r'\times')
        self.assertEqual(self.converter.OPERATORS['÷'], r'\div')
        self.assertEqual(self.converter.OPERATORS['·'], r'\cdot')
        self.assertEqual(self.converter.OPERATORS['≠'], r'\neq')
        self.assertEqual(self.converter.OPERATORS['≤'], r'\leq')
        self.assertEqual(self.converter.OPERATORS['≥'], r'\geq')

    def test_simple_conversion(self):
        """测试简单公式转换"""
        # 测试简单上标
        result = self.converter.convert_simple('x^2')
        self.assertEqual(result, 'x^{2}')

        # 测试简单下标
        result = self.converter.convert_simple('x_1')
        self.assertEqual(result, 'x_{1}')

        # 测试分式
        result = self.converter.convert_simple('a/b')
        self.assertEqual(result, r'\frac{a}{b}')

    def test_fallback_conversion(self):
        """测试回退转换"""
        # 测试 XML 标签移除
        xml = '<root><child>text</child></root>'
        result = self.converter._fallback_conversion(xml)
        self.assertEqual(result.strip(), 'text')

    def test_identifiers(self):
        """测试标识符转换"""
        # 使用简单字符串而不是 mock 对象
        # 直接测试 GREEK_LETTERS 映射
        self.assertEqual(self.converter.GREEK_LETTERS.get('alpha'), r'\alpha')
        self.assertEqual(self.converter.GREEK_LETTERS.get('beta'), r'\beta')


class TestTableConverter(unittest.TestCase):
    """表格转换器测试"""

    def setUp(self):
        self.converter = TableConverter()

    def test_clean_cell_content(self):
        """测试单元格内容清理"""
        # 测试 & 转义
        content = 'A & B'
        result = self.converter._clean_cell_content(content)
        self.assertEqual(result, r'A \& B')

        # 测试 % 转义
        content = '100%'
        result = self.converter._clean_cell_content(content)
        self.assertEqual(result, r'100\%')

        # 测试下划线转义
        content = 'text_1'
        result = self.converter._clean_cell_content(content)
        self.assertEqual(result, r'text\_1')

    def test_special_characters(self):
        """测试特殊字符转义"""
        # 注意：转义按顺序进行，\ 后 { 和 } 也会被转义
        special_cases = [
            ('$', r'\$'),
            ('#', r'\#'),
            ('_', r'\_'),
            ('~', r'\textasciitilde{}'),
            ('^', r'\textasciicircum{}'),
            # \ 会在 { } 之前被替换，所以 \ 后会被处理
            ('\\', r'\textbackslash{}'),
        ]

        for char, expected in special_cases:
            result = self.converter._clean_cell_content(char)
            self.assertEqual(result, expected)


class TestSimpleFormulas(unittest.TestCase):
    """简单公式转换测试（不需要 Word 文档）"""

    def test_fraction_conversion(self):
        """测试分式转换"""
        converter = MathMLToLaTeX()

        # MathML 分式结构
        mathml = """
        <math xmlns="http://www.w3.org/1998/Math/MathML">
            <mfrac>
                <mrow>
                    <mi>a</mi>
                    <mo>+</mo>
                    <mi>b</mi>
                </mrow>
                <mrow>
                    <mi>c</mi>
                </mrow>
            </mfrac>
        </math>
        """

        result = converter.convert(mathml)
        self.assertIn(r'\frac', result)
        self.assertIn('a', result)
        self.assertIn('b', result)
        # 验证分式结构
        self.assertIn('{c}', result)

    def test_sqrt_conversion(self):
        """测试平方根转换"""
        converter = MathMLToLaTeX()

        mathml = """
        <math xmlns="http://www.w3.org/1998/Math/MathML">
            <msqrt>
                <mi>x</mi>
            </msqrt>
        </math>
        """

        result = converter.convert(mathml)
        self.assertIn(r'\sqrt', result)
        self.assertIn('x', result)
        # 验证根号结构
        self.assertIn('{x}', result)

    def test_superscript_conversion(self):
        """测试上标转换"""
        converter = MathMLToLaTeX()

        mathml = """
        <math xmlns="http://www.w3.org/1998/Math/MathML">
            <msup>
                <mi>x</mi>
                <mn>2</mn>
            </msup>
        </math>
        """

        result = converter.convert(mathml)
        self.assertIn('x', result)
        self.assertIn('2', result)

    def test_subscript_conversion(self):
        """测试下标转换"""
        converter = MathMLToLaTeX()

        mathml = """
        <math xmlns="http://www.w3.org/1998/Math/MathML">
            <msub>
                <mi>x</mi>
                <mn>1</mn>
            </msub>
        </math>
        """

        result = converter.convert(mathml)
        self.assertIn('x', result)
        self.assertIn('1', result)


def run_tests():
    """运行所有测试"""
    # 创建测试套件
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    # 添加测试
    suite.addTests(loader.loadTestsFromTestCase(TestMathMLToLaTeX))
    suite.addTests(loader.loadTestsFromTestCase(TestTableConverter))
    suite.addTests(loader.loadTestsFromTestCase(TestSimpleFormulas))

    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    return result.wasSuccessful()


if __name__ == '__main__':
    success = run_tests()
    sys.exit(0 if success else 1)