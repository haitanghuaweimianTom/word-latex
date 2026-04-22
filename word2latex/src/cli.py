"""
Word → LaTeX 命令行接口
提供用户交互和参数解析
"""

from __future__ import annotations

import argparse
import sys
import os
from pathlib import Path

from converter import Word2LaTeXConverter, ConversionOptions


def create_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        prog='word2latex',
        description='将 Word 文档 (.docx) 转换为 LaTeX 代码',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python -m word2latex document.docx
  python -m word2latex document.docx -o output.tex
  python -m word2latex document.docx --img-dir ./images --package-mode
        """
    )

    parser.add_argument(
        'input_file',
        type=str,
        help='Word 文档路径 (.docx)'
    )

    parser.add_argument(
        '-o', '--output',
        type=str,
        default=None,
        help='输出文件路径（默认: 输出到标准输出）'
    )

    parser.add_argument(
        '--img-dir',
        type=str,
        default='./figures',
        help='图片提取目录（默认: ./figures）'
    )

    parser.add_argument(
        '--package-mode',
        action='store_true',
        help='仅输出 Document Body（不含导言区）'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='输出转换详情'
    )

    parser.add_argument(
        '--version',
        action='version',
        version='word2latex 1.0.0',
        help='显示版本信息'
    )

    return parser


def validate_input_file(filepath: str) -> bool:
    """验证输入文件"""
    path = Path(filepath)

    if not path.exists():
        print(f"错误: 文件不存在 '{filepath}'", file=sys.stderr)
        return False

    if not path.suffix.lower() == '.docx':
        print(f"警告: 文件扩展名不是 .docx '{filepath}'", file=sys.stderr)

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

    # 创建转换选项
    options = ConversionOptions(
        img_dir=args.img_dir,
        package_mode=args.package_mode,
        verbose=args.verbose
    )

    if args.verbose:
        print(f"输入文件: {args.input_file}")
        print(f"图片目录: {options.img_dir}")
        print(f"输出模式: {'包模式' if options.package_mode else '完整文档'}")
        print("-" * 50)

    try:
        # 创建转换器
        converter = Word2LaTeXConverter(options=options)

        if args.verbose:
            print("正在转换...")

        # 执行转换
        latex_code = converter.convert(args.input_file)

        # 输出结果
        if args.output:
            output_path = Path(args.output)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(latex_code, encoding='utf-8')

            if args.verbose:
                print(f"已保存到: {args.output}")
        else:
            # 输出到标准输出
            print(latex_code)

        if args.verbose:
            print("-" * 50)
            print("转换完成!")

    except Exception as e:
        print(f"转换失败: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()