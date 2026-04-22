#!/usr/bin/env python3
"""
运行所有测试和示例转换
"""

import os
import sys
import subprocess

# 确保在正确目录
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 切换到项目根目录
os.chdir('..')


def run_command(cmd, description):
    """运行命令并显示结果"""
    print(f"\n{'='*60}")
    print(f"📋 {description}")
    print(f"{'='*60}")
    print(f"命令: {cmd}")
    print("-" * 60)

    result = subprocess.run(cmd, shell=True, capture_output=False)

    if result.returncode == 0:
        print(f"✅ 成功")
    else:
        print(f"❌ 失败 (exit code: {result.returncode})")

    return result.returncode == 0


def main():
    print("=" * 60)
    print("🧪 Word2LaTeX 测试套件")
    print("=" * 60)

    # 1. 创建测试文档
    print("\n\n📄 步骤 1: 创建测试样例文档")
    print("-" * 60)

    result = run_command(
        "python tests/create_samples.py",
        "创建测试样例 Word 文档"
    )

    if not result:
        print("创建测试文档失败，退出")
        sys.exit(1)

    # 2. 列出测试文档
    print("\n\n📁 步骤 2: 检查测试文档")
    print("-" * 60)

    for fname in os.listdir('tests'):
        if fname.endswith('.docx'):
            fpath = os.path.join('tests', fname)
            size = os.path.getsize(fpath)
            print(f"  ✅ {fname} ({size:,} bytes)")

    # 3. 运行单元测试
    print("\n\n🧪 步骤 3: 运行单元测试")
    print("-" * 60)

    result = run_command(
        "python -m pytest tests/test_converter.py -v",
        "执行单元测试"
    )

    # 如果 pytest 不可用，使用 unittest
    if not result:
        print("\n尝试使用 unittest...")
        result = run_command(
            "python tests/test_converter.py",
            "执行单元测试 (unittest)"
        )

    # 4. 转换测试文档
    print("\n\n🔄 步骤 4: 转换测试文档为 LaTeX")
    print("-" * 60)

    test_files = [
        'tests/table_test.docx',
        'tests/complex_test.docx',
        'tests/math_test.docx',
        'tests/formula_test.docx',
    ]

    for docx_file in test_files:
        if os.path.exists(docx_file):
            output_file = docx_file.replace('.docx', '.tex')

            cmd = f"python -m word2latex {docx_file} -o {output_file} -v"
            run_command(cmd, f"转换 {os.path.basename(docx_file)}")

            # 显示输出文件大小
            if os.path.exists(output_file):
                size = os.path.getsize(output_file)
                print(f"  输出: {output_file} ({size:,} bytes)")

    # 5. 显示 LaTeX 输出示例
    print("\n\n📝 步骤 5: LaTeX 输出示例")
    print("-" * 60)

    output_file = 'tests/table_test.tex'
    if os.path.exists(output_file):
        print(f"\n📄 {output_file} 内容:")
        print("-" * 40)
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            # 只显示前 50 行
            lines = content.split('\n')
            for line in lines[:50]:
                print(line)
            if len(lines) > 50:
                print(f"... (共 {len(lines)} 行)")
    else:
        print("⚠️  table_test.tex 不存在")

    # 6. 检查图片提取
    print("\n\n🖼️ 步骤 6: 检查图片提取")
    print("-" * 60)

    if os.path.exists('figures'):
        files = os.listdir('figures')
        if files:
            print(f"已提取 {len(files)} 张图片:")
            for f in files:
                print(f"  - {f}")
        else:
            print("没有图片需要提取")
    else:
        print("figures 目录不存在（正常，因为测试文档没有图片）")

    print("\n\n" + "=" * 60)
    print("✨ 测试完成!")
    print("=" * 60)


if __name__ == '__main__':
    main()