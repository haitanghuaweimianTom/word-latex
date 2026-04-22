"""
LaTeX → Word 命令行接口
提供用户交互和参数解析
"""

from __future__ import annotations

import argparse
import sys
import os
from pathlib import Path

from converter import LaTeXToWordConverter, ConversionOptions
from parser import parse_latex


def create_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        prog='latex2word',
        description='将 LaTeX 文件 (.tex) 转换为 Word 文档 (.docx)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python -m latex2word document.tex
  python -m latex2word document.tex -o output.docx
  python -m latex2word document.tex --package-mode -v
        """
    )

    parser.add_argument(
        'input_file',
        type=str,
        help='LaTeX 文件路径 (.tex)'
    )

    parser.add_argument(
        '-o', '--output',
        type=str,
        default=None,
        help='输出文件路径（默认: 输入文件名.docx）'
    )

    parser.add_argument(
        '--package-mode',
        action='store_true',
        help='仅转换 Document Body（不含导言区）'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='输出转换详情'
    )

    parser.add_argument(
        '--version',
        action='version',
        version='latex2word 1.0.0',
        help='显示版本信息'
    )

    return parser


def validate_input_file(filepath: str) -> bool:
    """验证输入文件"""
    path = Path(filepath)

    if not path.exists():
        print(f"错误: 文件不存在 '{filepath}'", file=sys.stderr)
        return False

    if path.suffix.lower() not in ['.tex', '.ltx']:
        print(f"警告: 文件扩展名不是 .tex '{filepath}'", file=sys.stderr)

    if not path.is_file():
        print(f"错误: '{filepath}' 不是文件", file=sys.stderr)
        return False

    return True


def main():
    """主入口函数"""
    parser = create_parser()
    args = parser.parse_args()

    # 验证输入文件
    if not validate_input_file(args.input_file):
        sys.exit(1)

    # 确定输出文件
    input_path = Path(args.input_file)
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix('.docx')

    if args.verbose:
        print(f"输入文件: {args.input_file}")
        print(f"输出文件: {output_path}")
        print(f"模式: {'包模式' if args.package_mode else '完整文档'}")
        print("-" * 50)

    try:
        # 读取 LaTeX 文件
        if args.verbose:
            print("正在读取文件...")

        with open(args.input_file, 'r', encoding='utf-8') as f:
            latex_content = f.read()

        # 解析 LaTeX
        if args.verbose:
            print("正在解析 LaTeX...")

        latex_doc = parse_latex(latex_content)

        # 转换为 Word
        if args.verbose:
            print("正在转换为 Word...")

        options = ConversionOptions(
            package_mode=args.package_mode,
            verbose=args.verbose
        )
        converter = LaTeXToWordConverter(options=options)
        doc = converter.convert(latex_doc)

        # 保存文件
        if args.verbose:
            print("正在保存...")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))

        if args.verbose:
            print("-" * 50)
            print(f"转换完成! 输出文件: {output_path}")

    except Exception as e:
        print(f"转换失败: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()