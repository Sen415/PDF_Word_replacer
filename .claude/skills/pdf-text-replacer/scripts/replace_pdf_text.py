# -*- coding: utf-8 -*-
"""
Replace searchable PDF text while preserving the matched text style.
Supports both single edits (CLI flags) and batch edits (--edits-json).

Examples:
  python replace_pdf_text.py -i input.pdf -o output.pdf --page 1 --find "张三院长：" --replace "李四科长：" --verify
  python replace_pdf_text.py -i input.pdf -o output.pdf --edits-json edits.json --verify
  python replace_pdf_text.py -i input.pdf -o output.pdf --excel 替换表.xlsx --verify
"""

import argparse
import json
import re
import sys
import traceback
from pathlib import Path
from typing import Optional

import pymupdf

# ─── Windows 终端 UTF-8 兼容 ───
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BLACK = (0, 0, 0)
RED   = (1, 0, 0)

# ═══════════════════════════════════════════════════════════
#  内置 CJK 字体映射（PyMuPDF built-in fonts）
# ═══════════════════════════════════════════════════════════

FONT_MAP = {
    # PyMuPDF 内置 CJK 字体
    "fang":   "china-t",   # 仿宋
    "song":   "china-s",   # 宋体
    "hei":    "china-ss",  # 黑体
    "kai":    "china-ts",  # 楷体
    "li":     "china-cs",  # 隶书
    # 别名 — 拼音
    "fangsong":     "china-t",
    "simsun":       "china-s",
    "simhei":       "china-ss",
    "kaiti":        "china-ts",
    "lishu":        "china-cs",
    # 别名 — 中文全称
    "仿宋":     "china-t",
    "宋体":     "china-s",
    "黑体":     "china-ss",
    "楷体":     "china-ts",
    "隶书":     "china-cs",
    # 别名 — GB2312 / 新宋体 / 等线 / 微软雅黑 → 回退到对应风格
    "fangsonggb2312": "china-t",
    "kaitigb2312":    "china-ts",
    "nsimsun":        "china-s",
    "newsong":        "china-s",
    "dengxian":       "china-ss",
    "youyuan":        "china-t",
    "微软雅黑":        "china-ss",
    "microsoftyahei": "china-ss",
    "等线":            "china-ss",
    "新宋体":          "china-s",
    "幼圆":            "china-t",
}

# ═══════════════════════════════════════════════════════════
#  常见中文字体文件映射（用于尽量复用原 PDF 的真实字体）
#  Windows + macOS 双平台路径
# ═══════════════════════════════════════════════════════════

FONT_FILE_MAP = {
    # 仿宋
    "fangsong": [
        r"C:\Windows\Fonts\simfang.ttf",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
    ],
    "fangsonggb2312": [
        r"C:\Windows\Fonts\simfang.ttf",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
    ],
    # 宋体
    "simsun": [
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\simsun.ttf",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
    ],
    "songti": [
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\simsun.ttf",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
    ],
    "sunserif": [
        r"C:\Windows\Fonts\simsun.ttc",
    ],
    "nsimsun": [
        r"C:\Windows\Fonts\simsun.ttc",
    ],
    # 黑体
    "simhei": [
        r"C:\Windows\Fonts\simhei.ttf",
        "/System/Library/Fonts/STHeiti Light.ttc",
    ],
    "heiti": [
        r"C:\Windows\Fonts\simhei.ttf",
        "/System/Library/Fonts/STHeiti Light.ttc",
    ],
    "sans-serif": [
        r"C:\Windows\Fonts\simhei.ttf",
        "/System/Library/Fonts/STHeiti Light.ttc",
    ],
    "sansserif": [
        r"C:\Windows\Fonts\simhei.ttf",
        "/System/Library/Fonts/STHeiti Light.ttc",
    ],
    # 微软雅黑
    "microsoftyahei": [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
    ],
    "msyh": [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
    ],
    # 楷体
    "kaiti": [
        r"C:\Windows\Fonts\simkai.ttf",
        "/System/Library/Fonts/Supplemental/Kaiti.ttc",
    ],
    "kaitigb2312": [
        r"C:\Windows\Fonts\simkai.ttf",
        "/System/Library/Fonts/Supplemental/Kaiti.ttc",
    ],
    "serif": [
        r"C:\Windows\Fonts\simkai.ttf",
    ],
    # 隶书
    "lishu": [
        r"C:\Windows\Fonts\simli.ttf",
    ],
    # 等线
    "dengxian": [
        r"C:\Windows\Fonts\DengXian.ttf",
        r"C:\Windows\Fonts\Deng.ttf",
    ],
    "deng": [
        r"C:\Windows\Fonts\Deng.ttf",
    ],
    # 幼圆
    "youyuan": [
        r"C:\Windows\Fonts\SIMYOU.ttf",
    ],
    "simyou": [
        r"C:\Windows\Fonts\SIMYOU.ttf",
    ],
}

# ═══════════════════════════════════════════════════════════
#  字体回退映射（本地字体不存在 → PyMuPDF 内置字体）
# ═══════════════════════════════════════════════════════════

FONT_FALLBACK_MAP = {
    # 仿宋
    "fangsong":        "china-t",
    "fangsonggb2312":  "china-t",
    "华文仿宋":          "china-t",
    # 宋体
    "simsun":          "china-s",
    "songti":          "china-s",
    "sunserif":        "china-s",
    "nsimsun":         "china-s",
    "newsong":         "china-s",
    "新宋体":            "china-s",
    "华文宋体":          "china-s",
    # 黑体
    "simhei":          "china-ss",
    "heiti":           "china-ss",
    "sans-serif":      "china-ss",
    "sansserif":       "china-ss",
    "华文黑体":          "china-ss",
    # 微软雅黑 → 回退到黑体
    "microsoftyahei":  "china-ss",
    "msyh":            "china-ss",
    "微软雅黑":          "china-ss",
    "yahei":           "china-ss",
    # 楷体
    "kaiti":           "china-ts",
    "kaitigb2312":     "china-ts",
    "serif":           "china-ts",
    "华文楷体":          "china-ts",
    # 隶书
    "lishu":           "china-cs",
    "华文隶书":          "china-cs",
    # 等线 → 回退到黑体
    "dengxian":        "china-ss",
    "deng":            "china-ss",
    "等线":             "china-ss",
    # 幼圆 → 回退到仿宋
    "youyuan":         "china-t",
    "simyou":          "china-t",
    "幼圆":             "china-t",
}


# ═══════════════════════════════════════════════════════════
#  工具函数
# ═══════════════════════════════════════════════════════════

def normalize_font_key(fontname: str) -> str:
    """去掉 PDF 子集前缀和分隔符，方便匹配字体名"""
    if not fontname:
        return ""
    base = fontname.split("+")[-1]
    return base.lower().replace(" ", "").replace("-", "").replace("_", "")


def get_fontname(hint: Optional[str]) -> Optional[str]:
    """根据名称/别名查找内置 CJK 字体"""
    if not hint:
        return None
    key = hint.lower().replace(" ", "").replace("-", "").replace("_", "")
    return FONT_MAP.get(key, hint)


def rgb_from_pdf_int(value: int) -> tuple:
    """将 PyMuPDF span['color'] 的 0xRRGGBB 整数转为 0-1 RGB"""
    return (
        ((value >> 16) & 255) / 255,
        ((value >> 8) & 255) / 255,
        (value & 255) / 255,
    )


def parse_color(value) -> Optional[tuple]:
    """解析颜色：支持 '#RRGGBB'、'R,G,B'（0-255）、'R,G,B'（0-1）"""
    if value is None or value == "":
        return None

    value = str(value).strip()
    if value.startswith("#") and len(value) == 7:
        return (
            int(value[1:3], 16) / 255,
            int(value[3:5], 16) / 255,
            int(value[5:7], 16) / 255,
        )

    parts = [p.strip() for p in value.split(",")]
    if len(parts) != 3:
        raise ValueError("Color must be '#RRGGBB' or 'R,G,B'.")

    nums = [float(p) for p in parts]
    if any(n > 1 for n in nums):
        nums = [n / 255 for n in nums]
    return tuple(nums)


def find_local_font_file(fontname: str) -> Optional[str]:
    """根据 PDF 中的字体名寻找本机字体文件"""
    key = normalize_font_key(fontname)
    for alias, paths in FONT_FILE_MAP.items():
        if alias in key:
            for path in paths:
                if Path(path).exists():
                    return path
    return None


def fallback_fontname(fontname: str, default: str = "china-t") -> str:
    """原字体不可直接写入时，回退到 PyMuPDF 内置 CJK 字体"""
    key = normalize_font_key(fontname)
    for alias, mapped in FONT_FALLBACK_MAP.items():
        if alias in key:
            return mapped
    return default


# ═══════════════════════════════════════════════════════════
#  样式提取
# ═══════════════════════════════════════════════════════════

def get_original_text_style(page: pymupdf.Page, rect: pymupdf.Rect, find: str) -> dict:
    """从命中的原文字 span 中读取字体、字号、颜色和基线"""
    text_dict = page.get_text("dict")
    best = None

    for block in text_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")
                if not text.strip():
                    continue

                span_rect = pymupdf.Rect(span["bbox"])
                overlap = span_rect & rect
                if overlap.is_empty:
                    continue

                overlap_area = max(0, overlap.x1 - overlap.x0) * max(0, overlap.y1 - overlap.y0)
                score = (find in text, overlap_area)
                if best is None or score > best[0]:
                    best = (score, span)

    if best is None:
        return {}

    span = best[1]
    fontname = span.get("font", "")
    return {
        "fontsize": float(span.get("size", 12)),
        "fontname": fontname,
        "fontfile": find_local_font_file(fontname),
        "color": rgb_from_pdf_int(int(span.get("color", 0))),
        "origin": span.get("origin"),
    }


# ═══════════════════════════════════════════════════════════
#  核心替换逻辑
# ═══════════════════════════════════════════════════════════

def select_rects(rects: list, occurrence: Optional[int]) -> list:
    """按序号筛选匹配区域"""
    if occurrence is None:
        return rects
    if occurrence < 1:
        raise ValueError("--occurrence must be 1 or greater.")
    if occurrence > len(rects):
        return []
    return [rects[occurrence - 1]]


def insert_with_style(page: pymupdf.Page, rect: pymupdf.Rect, repl: str,
                      style: dict, fontname: Optional[str] = None,
                      fontsize: Optional[float] = None,
                      color: Optional[tuple] = None) -> None:
    """在擦除后的位置写入新文字，尽量保持原样式"""
    original_fontname = style.get("fontname")
    insert_fontfile = None if fontname else style.get("fontfile")
    insert_fontsize = fontsize or style.get("fontsize") or 12
    insert_color = color or style.get("color") or BLACK
    origin = style.get("origin")
    baseline_y = origin[1] if origin else rect.y1 - 2

    # 解析写入字体：优先用户指定 > 系统字体文件 > FONT_MAP 匹配 > 原文名称 > 仿宋
    if fontname:
        insert_fontname = fontname
    elif insert_fontfile:
        insert_fontname = original_fontname
    elif original_fontname:
        mapped = get_fontname(original_fontname)
        if mapped and mapped in ("china-t", "china-s", "china-ss", "china-ts", "china-cs"):
            insert_fontname = mapped
        else:
            insert_fontname = fallback_fontname(original_fontname, "china-t")
    else:
        insert_fontname = "china-t"

    try:
        page.insert_text(
            point=(rect.x0, baseline_y),
            text=repl,
            fontsize=insert_fontsize,
            fontname=insert_fontname,
            fontfile=insert_fontfile,
            color=insert_color,
        )
    except Exception:
        page.insert_text(
            point=(rect.x0, baseline_y),
            text=repl,
            fontsize=insert_fontsize,
            fontname=fallback_fontname(insert_fontname, "china-t"),
            color=insert_color,
        )


def replace_text(doc: pymupdf.Document, page_num: int, find: str,
                 repl: str, fontsize: Optional[float] = None,
                 fontname: Optional[str] = None,
                 color: Optional[tuple] = None,
                 occurrence: Optional[int] = None,
                 keep_style: bool = True) -> int:
    """在指定页面上搜索并替换文字。

    返回: 替换次数
    """
    page = doc[page_num]
    rects = select_rects(page.search_for(find), occurrence)
    if not rects:
        return 0

    matches = []
    for rect in rects:
        style = get_original_text_style(page, rect, find) if keep_style else {}
        matches.append((rect, style))
        page.add_redact_annot(rect)

    # ① 先统一擦除旧文字
    page.apply_redactions()

    # ② 逐处写入新文字
    for rect, style in matches:
        insert_with_style(page, rect, repl, style,
                          fontname=fontname, fontsize=fontsize, color=color)

    return len(matches)


# ═══════════════════════════════════════════════════════════
#  编辑列表加载
# ═══════════════════════════════════════════════════════════

def read_edits_from_args(args) -> list:
    """从命令行参数或 JSON/Excel 文件读取编辑列表"""
    if args.edits_json:
        with open(args.edits_json, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            raise ValueError("--edits-json must contain a list of edit objects.")
        return data

    if args.excel:
        return load_edits_from_excel(args.excel)

    missing = [name for name in ("page", "find", "replace")
               if getattr(args, name) in (None, "")]
    if missing:
        raise ValueError("Single edit mode requires --page, --find, and --replace.")
    return [{
        "page": args.page,
        "find": args.find,
        "replace": args.replace,
        "occurrence": args.occurrence,
        "fontsize": args.font_size,
        "fontname": args.font_name,
        "color": args.color,
    }]


def load_edits_from_excel(excel_path: str) -> list:
    """从 Excel 文件读取替换列表。

    Excel 列结构（第 1 行是表头）：
      A: 页码 (从1开始)
      B: 搜索文字
      C: 替换文字
      D: 字号 (可选，空白则继承原文字)
      E: 字体 (可选，空白则继承原文字；也可填 fang/song/hei/kai)
      F: 颜色 (可选，空白则继承原文字；格式 0,0,0)
    """
    try:
        import openpyxl
    except ImportError:
        print("需要 openpyxl 库: pip install openpyxl")
        sys.exit(1)

    wb = openpyxl.load_workbook(excel_path, read_only=True)
    ws = wb.active
    edits = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[1] or not row[2]:
            continue
        page = int(row[0]) - 1 if row[0] else 0
        find_txt = str(row[1]).strip()
        repl_txt = str(row[2]).strip()
        fsize = float(row[3]) if len(row) > 3 and row[3] else None
        fhint = str(row[4]).strip() if len(row) > 4 and row[4] else None
        color = None
        if len(row) > 5 and row[5]:
            parts = str(row[5]).split(",")
            if len(parts) == 3:
                color = tuple(float(x.strip()) / 255 for x in parts)

        edits.append({
            "page": page + 1,
            "find": find_txt,
            "replace": repl_txt,
            "fontsize": fsize,
            "fontname": fhint,
            "color": color,
        })

    wb.close()
    return edits


def normalize_edit(raw: dict) -> dict:
    """标准化单条编辑记录"""
    page = int(raw["page"])
    if page < 1:
        raise ValueError("Page numbers are 1-based. Use page=1 for the first page.")

    return {
        "page_num": page - 1,
        "find": str(raw["find"]),
        "replace": str(raw["replace"]),
        "occurrence": int(raw["occurrence"]) if raw.get("occurrence") else None,
        "fontsize": float(raw["fontsize"]) if raw.get("fontsize") else None,
        "fontname": get_fontname(raw.get("fontname")),
        "color": parse_color(raw.get("color")),
    }


# ═══════════════════════════════════════════════════════════
#  验证
# ═══════════════════════════════════════════════════════════

def verify_replacements(output_pdf: str, edits: list) -> None:
    """验证输出 PDF 中替换后的关键字是否存在"""
    doc = pymupdf.open(output_pdf)
    try:
        for edit in edits:
            page = doc[edit["page_num"]]
            rects = page.search_for(edit["replace"])
            status = "OK" if rects else "NOT FOUND"
            print("[VERIFY] Page %d: '%s' %s" %
                  (edit["page_num"] + 1, edit["replace"], status))
    finally:
        doc.close()


# ═══════════════════════════════════════════════════════════
#  主入口
# ═══════════════════════════════════════════════════════════

def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="Replace searchable PDF text while preserving original style.")
    parser.add_argument("--input", "-i", required=True, help="Input PDF path.")
    parser.add_argument("--output", "-o", required=True, help="Output PDF path.")

    # 单次替换
    parser.add_argument("--page", type=int,
                        help="1-based page number for single edit mode.")
    parser.add_argument("--find", help="Exact text to search for.")
    parser.add_argument("--replace", help="Replacement text.")

    # 批量替换
    parser.add_argument("--edits-json", help="JSON file containing a list of edit objects.")
    parser.add_argument("--excel", "-e",
                        help="Excel file with columns: 页码, 搜索文字, 替换文字, 字号, 字体, 颜色")

    parser.add_argument("--occurrence", type=int,
                        help="Only replace the Nth occurrence on the page.")
    parser.add_argument("--font-size", type=float,
                        help="Override font size. Omit to inherit.")
    parser.add_argument("--font-name",
                        help="Override font name (fang/song/hei/kai/li). Omit to inherit.")
    parser.add_argument("--color",
                        help="Override color as '#RRGGBB' or 'R,G,B'. Omit to inherit.")
    parser.add_argument("--no-keep-style", action="store_true",
                        help="Do not inherit original text style.")
    parser.add_argument("--verify", action="store_true",
                        help="Search output PDF for replacement text after saving.")
    parser.add_argument("--overwrite", action="store_true",
                        help="Allow replacing an existing output PDF.")
    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve()

    if input_path == output_path:
        raise ValueError("Output PDF must be different from input PDF.")
    if output_path.exists() and not args.overwrite:
        raise FileExistsError(
            "Output exists. Use --overwrite or choose a new output path: %s" % output_path)

    raw_edits = read_edits_from_args(args)
    edits = [normalize_edit(edit) for edit in raw_edits]
    doc = pymupdf.open(str(input_path))
    stats = {"total": 0, "ok": 0, "fail": 0}

    try:
        for edit in edits:
            try:
                count = replace_text(
                    doc,
                    edit["page_num"],
                    edit["find"],
                    edit["replace"],
                    fontsize=edit["fontsize"],
                    fontname=edit["fontname"],
                    color=edit["color"],
                    occurrence=edit["occurrence"],
                    keep_style=not args.no_keep_style,
                )
                if count:
                    stats["ok"] += 1
                    stats["total"] += count
                    print("[OK] Page %d: '%s' -> '%s' (%d)" %
                          (edit["page_num"] + 1, edit["find"], edit["replace"], count))
                else:
                    stats["fail"] += 1
                    print("[!!] Page %d: '%s' not found" %
                          (edit["page_num"] + 1, edit["find"]))
            except Exception:
                stats["fail"] += 1
                traceback.print_exc()

        output_path.parent.mkdir(parents=True, exist_ok=True)
        if hasattr(doc, "subset_fonts"):
            doc.subset_fonts()
        doc.save(str(output_path))
    finally:
        doc.close()

    print("[DONE] total=%d ok=%d fail=%d output=%s" %
          (stats["total"], stats["ok"], stats["fail"], output_path))

    if args.verify:
        verify_replacements(str(output_path), edits)


if __name__ == "__main__":
    main()
