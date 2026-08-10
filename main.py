# -*- coding: utf-8 -*-
"""
PDF 文字替换工具 — 命令行入口

用法:
  python main.py                            # 使用内置 demo 替换
  python main.py --excel 替换表.xlsx         # 从 Excel 读取替换列表
"""

import argparse

from pdf_replacer import (
    replace_from_list,
    load_edits_from_excel,
    verify,
)

# ══════════════════════════════════════════════
#  主程序 (single/Excel-driven replace)
# ══════════════════════════════════════════════

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF 文字替换工具")
    parser.add_argument("--excel",  "-e", help="从 Excel 文件读取替换列表")
    parser.add_argument("--input",  "-i", help="输入 PDF 路径")
    parser.add_argument("--output", "-o", help="输出 PDF 路径")
    parser.add_argument("--verify", "-v", action="store_true", help="替换后验证")
    args = parser.parse_args()

    # ── 默认 Demo ──
    input_pdf  = args.input  or "example_input.pdf"
    output_pdf = args.output or "example_output.pdf"

    if args.excel:
        edits = load_edits_from_excel(args.excel)
    else:
        edits = [
            (0, "张三", "李四", None, None, None),

        ]

    print("=" * 50)
    print("输入: %s" % input_pdf)
    print("输出: %s" % output_pdf)
    print("替换条目: %d" % len(edits))
    print("=" * 50)

    stats = replace_from_list(input_pdf, output_pdf, edits)

    print("=" * 50)
    print("完成 — 共 %d 处替换, %d 项成功, %d 项未找到" %
          (stats["total"], stats["ok"], stats["fail"]))
    print("保存至: %s" % output_pdf)

    if args.verify:
        checks = [(e[0], e[2]) for e in edits]
        print()
        verify(output_pdf, checks)
