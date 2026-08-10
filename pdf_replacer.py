# -*- coding: utf-8 -*-
"""
PDF 文字替换工具 — 核心模块

基于 PyMuPDF 实现，核心逻辑为"先擦后写"：
  1. page.search_for(text)   → 定位旧文字位置
  2. page.add_redact_annot() → 标记擦除区域
  3. page.apply_redactions() → 永久删除旧文字
  4. page.insert_text()      → 在原位写入新文字
"""

import sys
import traceback
from pathlib import Path
from typing import Optional

import pymupdf

# ─── Windows 终端 UTF-8 兼容 ───
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ─── 内置 CJK 字体映射 ───
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
    # 别名 — GB2312 / 新宋体 / 等线 → 回退到对应风格
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

# ─── 默认颜色 ───
BLACK = (0, 0, 0)
RED   = (1, 0, 0)


# ─── 常见中文字体文件映射：用于尽量复用原 PDF 的真实字体 ───
FONT_FILE_MAP = {
    # 仿宋
    "fangsong":        [r"C:\Windows\Fonts\simfang.ttf"],
    "fangsonggb2312":  [r"C:\Windows\Fonts\simfang.ttf"],
    # 宋体
    "simsun":          [r"C:\Windows\Fonts\simsun.ttc", r"C:\Windows\Fonts\simsun.ttf"],
    "songti":          [r"C:\Windows\Fonts\simsun.ttc", r"C:\Windows\Fonts\simsun.ttf"],
    "sunserif":        [r"C:\Windows\Fonts\simsun.ttc"],
    "nsimsun":         [r"C:\Windows\Fonts\simsun.ttc"],
    # 黑体
    "simhei":          [r"C:\Windows\Fonts\simhei.ttf"],
    "heiti":           [r"C:\Windows\Fonts\simhei.ttf"],
    "sans-serif":      [r"C:\Windows\Fonts\simhei.ttf"],
    "sansserif":       [r"C:\Windows\Fonts\simhei.ttf"],
    # 微软雅黑
    "microsoftyahei":  [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyh.ttf"],
    "msyh":            [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyh.ttf"],
    # 楷体
    "kaiti":           [r"C:\Windows\Fonts\simkai.ttf"],
    "kaitigb2312":     [r"C:\Windows\Fonts\simkai.ttf"],
    "serif":           [r"C:\Windows\Fonts\simkai.ttf"],
    # 隶书
    "lishu":           [r"C:\Windows\Fonts\simli.ttf"],
    # 等线
    "dengxian":        [r"C:\Windows\Fonts\DengXian.ttf", r"C:\Windows\Fonts\Deng.ttf"],
    "deng":            [r"C:\Windows\Fonts\Deng.ttf"],
    # 幼圆
    "youyuan":         [r"C:\Windows\Fonts\SIMYOU.ttf"],
    "simyou":          [r"C:\Windows\Fonts\SIMYOU.ttf"],
}

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


def normalize_font_key(fontname: str) -> str:
    """去掉 PDF 子集前缀和分隔符，方便匹配字体名"""
    if not fontname:
        return ""
    base = fontname.split("+")[-1]
    return base.lower().replace(" ", "").replace("-", "").replace("_", "")


def rgb_from_pdf_int(value: int) -> tuple:
    """将 PyMuPDF span['color'] 的 0xRRGGBB 整数转为 0-1 RGB"""
    return (
        ((value >> 16) & 255) / 255,
        ((value >> 8) & 255) / 255,
        (value & 255) / 255,
    )


def find_local_font_file(fontname: str) -> Optional[str]:
    """根据 PDF 中的字体名寻找本机字体文件"""
    key = normalize_font_key(fontname)
    for alias, paths in FONT_FILE_MAP.items():
        if alias in key:
            for path in paths:
                if Path(path).exists():
                    return path
    return None


def fallback_fontname(fontname: str, default: str) -> str:
    """原字体不可直接写入时，回退到 PyMuPDF 内置 CJK 字体"""
    key = normalize_font_key(fontname)
    for alias, mapped in FONT_FALLBACK_MAP.items():
        if alias in key:
            return mapped
    return default


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


def get_fontname(hint: Optional[str]) -> Optional[str]:
    """根据名称/别名查找内置 CJK 字体"""
    if not hint:
        return None
    key = hint.lower().replace(" ", "").replace("-", "").replace("_", "")
    return FONT_MAP.get(key, hint)


def replace_text(
    doc: pymupdf.Document,
    page_num: int,
    find: str,
    repl: str,
    fontsize: Optional[float] = None,
    fontname: Optional[str] = None,
    color: Optional[tuple] = None,
    keep_original_style: bool = True,
) -> int:
    """在指定页面上搜索并替换文字。

    返回: 替换次数
    """
    page = doc[page_num]
    areas = page.search_for(find)

    if not areas:
        return 0

    matches = []
    for rect in areas:
        style = get_original_text_style(page, rect, find) if keep_original_style else {}
        matches.append((rect, style))
        page.add_redact_annot(rect)

    # ① 先统一擦除旧文字
    page.apply_redactions()

    for rect, style in matches:
        original_fontname = style.get("fontname")
        insert_fontfile = None if fontname else style.get("fontfile")
        insert_fontsize = fontsize or style.get("fontsize") or 12
        insert_color = color or style.get("color") or BLACK
        origin = style.get("origin")
        baseline_y = origin[1] if origin else rect.y1 - 2

        # ② 解析写入字体：优先用户指定 > 系统字体文件 > FONT_MAP 匹配 > 原文名称 > 仿宋
        if fontname:
            insert_fontname = fontname
        elif insert_fontfile:
            insert_fontname = original_fontname  # PyMuPDF 要求 fontname 不能为 None，否则 .startswith() 崩溃
        elif original_fontname:
            mapped = get_fontname(original_fontname)
            if mapped and mapped in ("china-t", "china-s", "china-ss", "china-ts", "china-cs"):
                insert_fontname = mapped
            else:
                insert_fontname = fallback_fontname(original_fontname, "china-t")
        else:
            insert_fontname = "china-t"

        # ③ 写入新文字
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

    return len(areas)


def replace_from_list(
    input_pdf: str,
    output_pdf: str,
    edits: list,
) -> dict:
    """Replace text in a PDF from a list of (page, find, repl, ...) tuples.

    Returns a dict with keys: total, ok, fail.
    """
    doc = pymupdf.open(input_pdf)
    stats = {"total": 0, "ok": 0, "fail": 0}

    for entry in edits:
        page_num, find_text, repl_text, fontsize, font_hint, color = entry
        fontname = get_fontname(font_hint)

        try:
            n = replace_text(doc, page_num, find_text, repl_text, fontsize, fontname, color)
            if n > 0:
                print("  %s  →  %s  (%d处)" %
                      (find_text, repl_text, n))
                stats["ok"] += 1
                stats["total"] += n
            else:
                print("[!!] Page %d: '%s' 未找到" % (page_num + 1, find_text))
                stats["fail"] += 1
        except Exception:
            print("[!!] Page %d: '%s' 替换失败" % (page_num + 1, find_text))
            traceback.print_exc()
            stats["fail"] += 1

    doc.subset_fonts()
    doc.save(output_pdf)
    doc.close()

    return stats


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
        page     = int(row[0]) - 1 if row[0] else 0
        find_txt = str(row[1]).strip()
        repl_txt = str(row[2]).strip()
        fsize    = float(row[3]) if len(row) > 3 and row[3] else None
        fhint    = str(row[4]).strip() if len(row) > 4 and row[4] else None
        color    = None
        if len(row) > 5 and row[5]:
            parts = str(row[5]).split(",")
            if len(parts) == 3:
                color = tuple(float(x.strip()) / 255 for x in parts)

        edits.append((page, find_txt, repl_txt, fsize, fhint, color))

    wb.close()
    return edits


def verify(output_pdf: str, checks: list) -> None:
    """验证输出 PDF 中关键字是否存在"""
    doc = pymupdf.open(output_pdf)
    for page_num, keyword in checks:
        page = doc[page_num]
        r = page.search_for(keyword)
        status = "OK" if r else "NOT FOUND"
        print("[VERIFY] Page %d: '%s' %s" % (page_num + 1, keyword, status))
    doc.close()
