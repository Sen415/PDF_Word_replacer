# -*- coding: utf-8 -*-
"""批量 PDF 文字替换工具 — 命令行入口

一份模板 + 一张数据表 → 多份个性化 PDF。

══════════════════════════════════════════════
  用法一：在下方配置区填写参数，直接运行
    python batch_invitation.py

  用法二：通过命令行参数运行
    python batch_invitation.py -i 模板.pdf -e 数据.xlsx -d 输出/ \\
        --find "尊敬的来宾：" --replace "尊敬的{{姓名}}{{职务}}："
══════════════════════════════════════════════
"""

import argparse
from pathlib import Path

from batch_replacer import batch_replace, derive_filename_suffix


# ═══════════════════════════════════════════════════
#  【配置区】直接修改下方变量即可运行
#  (命令行传参会覆盖这里的默认值)
# ═══════════════════════════════════════════════════

SCRIPT_DIR = Path(__file__).resolve().parent

DEFAULTS = {
    "input_pdf":    str(SCRIPT_DIR / "example_input.pdf"),       # ← 模板 PDF 路径
    "excel":        str(SCRIPT_DIR / "人员职务名单.xlsx"),      # ← 数据 Excel 路径
    "output_dir":   str(SCRIPT_DIR / "批量邀请函"),      # ← 输出目录
    "find":         "张三院长：",           # ← 模板中要查找并替换掉的文字
    "replace":      "{{姓名}}{{职务|short}}：",  # ← 替换为（{{列名}} / {{列名|short}}）
    "filename":     "",                  # ← 输出文件名模板（留空则自动：{{姓名}}.pdf）
    "filename_suffix_from": "",          # ← 从模板 PDF 文件名推导后缀（与 --find 配合）
    "sheet":        "",                  # ← Excel 工作表名（留空则自动发现）
    "page":         1,                   # ← 替换在 PDF 第几页（从 1 开始）
    "verify":       False,               # ← 生成后是否验证
    "overwrite":    False,               # ← 是否覆盖已存在文件
    "dry_run":      False,               # ← True=只预览不生成 / False=实际生成
}


# ═══════════════════════════════════════════════════
#  命令行（通常不需要修改下方代码）
# ═══════════════════════════════════════════════════

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="批量 PDF 文字替换工具 — 一份模板 + 一张数据表 → 多份个性化 PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
模板语法:
  {{列名}}          取该列的值
  {{列名|short}}    取该列的值并提取短称谓（如 设备科科长 → 科长）

示例:
  python batch_invitation.py -i 模板.pdf -e 人员表.xlsx -d 输出/ \\
      --find "尊敬的来宾：" --replace "尊敬的{{姓名}}{{职务}}："
  python batch_invitation.py -i 模板.pdf -e 数据.xlsx -d 输出/ \\
      --find "XXX" --replace "{{姓名}}" --dry-run
        """,
    )

    parser.add_argument(
        "--input", "-i", default=DEFAULTS["input_pdf"],
        help="模板 PDF 路径（默认: %(default)s）")
    parser.add_argument(
        "--excel", "-e", default=DEFAULTS["excel"],
        help="数据 Excel 路径（默认: %(default)s）")
    parser.add_argument(
        "--output-dir", "-d", default=DEFAULTS["output_dir"],
        help="输出目录（默认: %(default)s）")
    parser.add_argument(
        "--find", "-f", default=DEFAULTS["find"],
        help="查找的固定文字（默认: %(default)s）")
    parser.add_argument(
        "--replace", "-r", default=DEFAULTS["replace"],
        help="替换模板，支持 {{列名}} / {{列名|short}}（默认: %(default)s）")
    parser.add_argument(
        "--filename", "-n", default=DEFAULTS["filename"] or None,
        help="输出文件名模板。留空则自动: {{姓名}}.pdf")
    parser.add_argument(
        "--filename-suffix-from", metavar="PDF",
        default=DEFAULTS["filename_suffix_from"] or None,
        help="从指定 PDF 文件名推导输出文件名后缀")
    parser.add_argument(
        "--sheet", "-s", default=DEFAULTS["sheet"] or None,
        help="Excel 工作表名（留空则自动发现）")
    parser.add_argument(
        "--page", "-p", type=int, default=DEFAULTS["page"],
        help="替换目标页码，从 1 开始（默认: %(default)s）")
    parser.add_argument(
        "--verify", "-v", action="store_true", default=DEFAULTS["verify"],
        help="生成后验证替换文字是否存在")
    parser.add_argument(
        "--overwrite", action="store_true", default=DEFAULTS["overwrite"],
        help="覆盖已存在的同名文件")
    parser.add_argument(
        "--dry-run", action="store_true", default=DEFAULTS["dry_run"],
        help="预览模式：只显示将要执行的操作，不实际生成文件")

    args = parser.parse_args()

    # ── 页码转换：用户填的是从 1 开始，内部是 0-based ──
    page_0based = args.page - 1

    # ── 处理文件名模板 ──
    filename_template = args.filename
    if not filename_template and args.filename_suffix_from:
        suffix = derive_filename_suffix(args.filename_suffix_from, args.find)
        filename_template = "{{姓名}}{{职务|short}}" + suffix
        print("从模板文件名推导后缀: %s" % suffix)
    elif not filename_template:
        filename_template = "{{姓名}}.pdf"

    print("=" * 50)
    print("模板 PDF : %s" % args.input)
    print("数据 Excel: %s" % args.excel)
    print("输出目录  : %s" % args.output_dir)
    print("查找文字  : %s" % args.find)
    print("替换模板  : %s" % args.replace)
    print("文件名    : %s" % filename_template)
    print("目标页码  : 第 %d 页" % args.page)
    if args.dry_run:
        print("*** 预览模式（不生成文件）***")
    print("=" * 50)

    stats = batch_replace(
        input_pdf=args.input,
        excel_path=args.excel,
        output_dir=args.output_dir,
        find_text=args.find,
        replace_template=args.replace,
        filename_template=filename_template,
        sheet_name=args.sheet,
        page_num=page_0based,
        verify_result=args.verify,
        overwrite=args.overwrite,
        dry_run=args.dry_run,
    )

    print("=" * 50)
    if args.dry_run:
        print("预览完成 — 共 %d 行数据，将生成 %d 份 PDF" %
              (stats["total_rows"], stats["ok"]))
    else:
        print("完成 — 数据 %d 行, 生成 %d 份, 跳过 %d 份, 失败 %d 份" %
              (stats["total_rows"], stats["ok"], stats["skipped"], stats["fail"]))
        if stats["outputs"]:
            print("保存目录: %s" % args.output_dir)
