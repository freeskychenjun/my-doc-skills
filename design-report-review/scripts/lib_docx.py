# -*- coding: utf-8 -*-
"""
lib_docx.py — 报告智能校审技能 · docx 解析基础库

提供：
  1) 文档顺序遍历（段落与表格交错，保留原始顺序，便于题录/引用/数据定位）
  2) 表格网格重建（正确处理横向合并 gridSpan 与纵向合并 vMerge）
  3) .doc → .docx 的尽力转换
  4) 检查所需的正则常量（题录 / 标准引用 / 法规引用 / 数值 / 罗马数字）

仅依赖 python-docx 与标准库。所有函数对中英文标点全/半角均做归一化处理。
"""

from __future__ import annotations
import os
import re
import sys
import unicodedata

# ---------- oxml 依赖（惰性导入，便于单测） ----------
from docx.document import Document as _DocumentClass
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


# ============================================================
# 1. 文档顺序遍历
# ============================================================
def iter_block_items(parent):
    """按文档原始顺序产出 ('paragraph', Paragraph) 或 ('table', Table)。

    python-docx 的 .paragraphs / .tables 是两个独立列表，丢失了相对顺序；
    本函数直接遍历 body 的子元素，保证段落与表格交错顺序正确。
    """
    if isinstance(parent, _DocumentClass):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        # 容错：尝试取 element.body
        parent_elm = getattr(parent, "element", parent).body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield ("paragraph", Paragraph(child, parent))
        elif isinstance(child, CT_Tbl):
            yield ("table", Table(child, parent))


def para_text(p: Paragraph) -> str:
    """段落文本。优先用 Paragraph.text；为空时回退到 w:t 拼接（兼容域代码）。"""
    try:
        t = p.text
    except Exception:
        t = ""
    if t and t.strip():
        return t
    # 回退：直接取所有 w:t
    parts = []
    for node in p._p.iter(qn("w:t")):
        parts.append(node.text or "")
    return "".join(parts)


def para_style(p: Paragraph) -> str:
    try:
        return p.style.name if p.style else ""
    except Exception:
        return ""


# ============================================================
# 2. 表格网格重建（处理合并单元格）
# ============================================================
def _tc_text(tc) -> str:
    """从 w:tc 元素提取文本，段落以换行连接。"""
    paras = []
    for p in tc.iter(qn("w:p")):
        runs = []
        for t in p.iter(qn("w:t")):
            runs.append(t.text or "")
        paras.append("".join(runs))
    return "\n".join(s for s in paras if s != "")


def grid_from_table(tbl: Table):
    """将表格重建为二维文本网格 [[str|None, ...], ...]。

    - 横向合并(gridSpan)：首列放文本，被合并列置 None。
    - 纵向合并(vMerge restart)：起点放文本；后续 continue 行对应列置 None。
    这样行列合计等算术检查不会因合并单元格重复计数。
    """
    tbl_el = tbl._tbl
    rows_xml = tbl_el.findall(qn("w:tr"))
    grid = []          # grid[r][c] = str | None
    col_count = 0
    for r, tr in enumerate(rows_xml):
        if len(grid) <= r:
            grid.append([])
        row = grid[r]
        c = 0
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            grid_span = 1
            v_merge = None  # None | 'restart' | 'continue'
            if tcPr is not None:
                gs = tcPr.find(qn("w:gridSpan"))
                if gs is not None:
                    try:
                        grid_span = int(gs.get(qn("w:val")) or "1")
                    except ValueError:
                        grid_span = 1
                vm = tcPr.find(qn("w:vMerge"))
                if vm is not None:
                    val = vm.get(qn("w:val"))
                    v_merge = "restart" if val == "restart" else "continue"
            text = _tc_text(tc).strip()
            cols = list(range(c, c + grid_span))
            # 扩展本行长度
            need = c + grid_span
            while len(row) < need:
                row.append(None)
            if v_merge == "continue":
                # 续接上方合并：本格留空（文本归属起点行）
                for col in cols:
                    row[col] = None
            else:
                # 普通格或 vMerge 起点：首列写文本，其余合并列置 None
                row[c] = text
                for col in cols[1:]:
                    row[col] = None
            c += grid_span
        col_count = max(col_count, len(row))
    # 各行对齐到 col_count
    for row in grid:
        while len(row) < col_count:
            row.append(None)
    return grid


# ============================================================
# 3. .doc → .docx 转换（尽力）
# ============================================================
def ensure_docx(path: str) -> str:
    """若为 .doc，尝试转为 .docx 并返回新路径；.docx 原样返回。"""
    p = os.path.abspath(path)
    low = p.lower()
    if low.endswith(".docx"):
        return p
    if not low.endswith(".doc"):
        # 既非 doc 也非 docx，仍尝试当作 docx 读取
        return p

    out_dir = os.path.dirname(p) or "."
    base = os.path.splitext(os.path.basename(p))[0]
    target = os.path.join(out_dir, base + "_converted.docx")

    # 方案1：LibreOffice / soffice
    for exe in ("soffice", "libreoffice", "soffice.exe"):
        import shutil
        if shutil.which(exe):
            import subprocess
            try:
                subprocess.run(
                    [exe, "--headless", "--convert-to", "docx",
                     "--outdir", out_dir, p],
                    check=True, timeout=300,
                )
                cand = os.path.join(out_dir, base + ".docx")
                if os.path.exists(cand):
                    return cand
            except Exception:
                pass

    # 方案2：Windows Word COM (pywin32)
    word = None
    try:
        import win32com.client  # type: ignore
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        d = word.Documents.Open(p, ReadOnly=True)
        d.SaveAs(target, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
        d.Close(False)
        # SaveAs 成功后立即检查文件——不因后续 Quit 的 RPC 偶发错误而判失败
        if os.path.exists(target):
            return target
    except Exception:
        pass
    finally:
        # word.Quit() 单独兜底：Word 进程退出时偶发 RPC 错误，不影响已生成的文件
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

    # 最后再确认一次（Quit 异常情况下文件可能已生成）
    if os.path.exists(target):
        return target

    raise RuntimeError(
        "无法转换 .doc 文件为 .docx。请先用 Word 另存为 .docx，"
        "或安装 LibreOffice（soffice）/ pywin32。原文件：" + p
    )


# ============================================================
# 4. 归一化与正则常量
# ============================================================
DASH_SET = "—－-‐–‒—·．."  # 全/半角破折号、连字符、中点等

def norm_code(code: str) -> str:
    """标准化标准/法规编号，便于比对：去空格、统一分隔符为 '-'、大写。

    'SL/T 252-2017' -> 'SL/T252-2017'
    'GB50265—2022'  -> 'GB50265-2022'
    'SL/T791—2019'  -> 'SL/T791-2019'
    """
    if not code:
        return ""
    s = code.strip()
    # 去除参考清单中偶发的脏字符（如 "SL/'T 292" 里的引号/撇号）
    s = re.sub(r"['’“”\"`]", "", s)
    # 全角斜杠转半角
    s = s.replace("／", "/")
    # 把所有形式的破折号/连接符统一为 '-'
    s = re.sub(r"[—－‐–‒﹣\-．.]", "-", s)
    s = re.sub(r"\s+", "", s)
    s = s.upper()
    return s


def norm_name(name: str) -> str:
    """标准化名称：去除版本后缀与空白，便于按名称匹配。"""
    if not name:
        return ""
    s = name.strip().strip("《》")
    # 去除括注后缀：(修订)/(修正)/(制定)/(年)/(版) 等
    s = re.sub(r"[（(][^（()]*[）)]", "", s)
    s = re.sub(r"\s+", "", s)
    return s


# 标准编号：前缀(SL/GB/NB/DL/JTG/JTS/CJJ/DB) + 可选 /T 或 /Z + 编号 + 可选年份
STD_CODE_RE = re.compile(
    r"(?<![\w/])"
    r"(SL|GB|NB|DL|JTG|JTS|CJJ|DB)"
    r"\s*(/(?:T|Z))?"
    r"\s*(\d+(?:\.\d+)?)"
    r"(?:\s*([—\-–．.‐-―·])\s*(\d{4}))?"
    r"(?!\d)"
)

# 法规/规章令号：主席令 / 国务院令 / 部令 / 文号 等
LAW_CODE_RE = re.compile(
    r"(\d{2,4}\s*年)?\s*"
    r"(主席令|国务院令|国务院第[\d〇零一二三四五六七八九十百千]+号令|"
    r"国务院第[\d〇零一二三四五六七八九十百千]+号|"
    r"[一-龥]{2,8}令第[\d〇零一二三四五六七八九十百千]+号|"
    r"[一-龥]{2,8}第[\d〇零一二三四五六七八九十百千]+号令|"
    r"国函〔?\d{4}〕?\s*第?\d+号|"
    r"[一-龥]{1,8}〔\d{4}〕\s*第?\d+号)"
)

# 书名号引用《...》，候选法规/标准名称
GUIBOOK_RE = re.compile(r"《([^《》]{2,60})》")

# 图/表题录：图2.1-1 / 表 3.2-3 等
CAPTION_RE = re.compile(
    r"(图|表)\s*([0-9]+(?:[\.\-][0-9]+){1,3})\s*([^\n，。；！？]*)"
)

# 仅"以图/表开头"的题录定义行（更可靠，用于编号顺序检查）
CAPTION_DEF_RE = re.compile(r"^\s*(图|表)\s*([0-9]+(?:[\.\-][0-9]+){1,3})\b")

# 流量值：1290m³/s / 5 m3/s / 256.34m³/s
FLOW_RE = re.compile(r"([0-9]+(?:\.[0-9]+)?)\s*[～\-~至到]?[ ]?(?:[0-9]+(?:\.[0-9]+)?)?\s*m[³3]\s*/\s*s")

# 排涝模数：0.50～0.65m³/s·km² / 0.5m3/s/km2 等多种写法（中点或斜杠分隔）
MODULUS_RE = re.compile(
    r"([0-9]+(?:\.[0-9]+)?)\s*(?:[～\-~至到]\s*([0-9]+(?:\.[0-9]+)?)\s*)?"
    r"m\s*[³3]\s*/\s*s\s*[·•・.·/\s]*\s*km\s*[²2]"
)

# 面积
AREA_KM2_RE = re.compile(r"([0-9]+(?:\.[0-9]+)?)\s*km\s*[²2]")
AREA_MU_RE = re.compile(r"([0-9]+(?:\.[0-9]+)?)\s*万\s*亩")

# 大数单位（std_001）：数字+亿/万/百万；以及 10的N次方/10^N/10ⁿ
BIGUNIT_RE = re.compile(r"([0-9]+(?:\.[0-9]+)?)\s*(亿|万|百万)")
POW10_RE = re.compile(r"10\s*的?\s*([0-9]+)\s*次方|10\s*\^\s*([0-9]+)|10\s*([⁰¹²³⁴⁵⁶⁷⁸⁹]+)")

# 罗马数字（含全角）与阿拉伯 1-7
ROMAN_CHARS = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ" + "IIIIIIVVVIVIIVIIIIXXXIXII"
ROMAN_RE = re.compile(r"[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]")

# 工程等别/级别上下文（std_004）
GRADE_CTX_RE = re.compile(
    r"(工程等别|工程级别|建筑物级别|水工建筑物级别|等别|级别|等级|航道等级|船闸等级|围岩类别|可钻性|阶地级别|地面水环境)"
)


def is_roman(s: str) -> bool:
    return bool(s) and all(ch in ROMAN_CHARS for ch in s if not ch.isspace())


def to_arabic(roman: str) -> int:
    """罗马数字（含全角）转阿拉伯，仅支持 I..XII 常见工程等别范围。"""
    m = {"Ⅰ": 1, "II": 1, "Ⅱ": 2, "III": 3, "Ⅲ": 3, "IV": 4, "Ⅳ": 4,
         "V": 5, "Ⅴ": 5, "VI": 6, "Ⅵ": 6, "VII": 7, "Ⅶ": 7, "VIII": 8,
         "Ⅷ": 8, "IX": 9, "Ⅸ": 9, "X": 10, "Ⅹ": 10, "XI": 11, "Ⅺ": 11,
         "XII": 12, "Ⅻ": 12}
    return m.get(roman.strip(), -1)


__all__ = [
    "iter_block_items", "para_text", "para_style",
    "grid_from_table", "ensure_docx",
    "norm_code", "norm_name", "is_roman", "to_arabic",
    "STD_CODE_RE", "LAW_CODE_RE", "GUIBOOK_RE",
    "CAPTION_RE", "CAPTION_DEF_RE",
    "FLOW_RE", "MODULUS_RE", "AREA_KM2_RE", "AREA_MU_RE",
    "BIGUNIT_RE", "POW10_RE", "ROMAN_RE", "GRADE_CTX_RE",
    "DASH_SET",
]
