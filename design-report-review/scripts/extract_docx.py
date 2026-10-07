# -*- coding: utf-8 -*-
"""
extract_docx.py — 报告智能校审技能 · 文档结构化抽取器

输入：docx/doc 路径
输出目录（默认 <skill>/_run/<文档名>/）下生成：
  fulltext.md     带段落序号与标题层级的全文（表格只留标记，详细数据见 tables.json）
  tables.json     全部表格（2D 网格 + 邻近题录 + 所在章节 + 表序）
  captions.json   图/表题录（类型、编号、文本、是否为定义行、段落序号、章节）
  citations.json  标准与法规引用（原始串、前缀、号、年、归一化码、邻近《》名称、段落序号、章节）
  numbers.json    关键数值（流量/排涝模数/面积/大数单位/10的次方）及其上下文与定位
  meta.json       文档元信息（段落数、表数、标题树、章节列表、文件名）

用法：
  python extract_docx.py <docx路径> [输出目录]
"""

from __future__ import annotations
import json
import os
import re
import sys

# GBK 控制台自愈：打印 ✓/✗ 等会 UnicodeEncodeError 且产物不落盘（2026-09-30 实测）
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib_docx import (  # noqa: E402
    iter_block_items, para_text, para_style, grid_from_table, ensure_docx,
    STD_CODE_RE, LAW_CODE_RE, GUIBOOK_RE, CAPTION_RE, CAPTION_DEF_RE,
    FLOW_RE, MODULUS_RE, AREA_KM2_RE, AREA_MU_RE, BIGUNIT_RE, POW10_RE,
    norm_code, norm_name,
)
from docx import Document  # noqa: E402


HEADING_RE = re.compile(r"(?:Heading|标题|heading|Title)\s*([1-9])", re.IGNORECASE)


def heading_level(style_name: str):
    if not style_name:
        return None
    m = HEADING_RE.search(style_name)
    if m:
        return int(m.group(1))
    # 纯 "Title" / "标题" 视为 1
    if style_name.lower() in ("title", "标题"):
        return 1
    return None


def _context(text, start, end, win=35):
    a = max(0, start - win)
    b = min(len(text), end + win)
    return text[a:b].replace("\n", " ").strip()


def extract(docx_path: str, out_dir: str):
    docx_path = ensure_docx(docx_path)
    doc = Document(docx_path)

    fulltext_lines = []
    tables, captions, citations, numbers = [], [], [], []
    headings = []          # {level, text, para_idx}
    sections = []          # {level, text, para_idx}
    section_stack = {}     # level -> text
    heading_by_para = {}   # para_idx -> 最近标题文本(用于定位)
    last_heading = ""

    para_idx = -1
    tbl_idx = -1

    for kind, obj in iter_block_items(doc):
        if kind == "paragraph":
            para_idx += 1
            style = para_style(obj)
            text = para_text(obj)
            lvl = heading_level(style)
            tag = f"[H{lvl}]" if lvl else ""
            pid = f"P{para_idx:04d}"

            if lvl:
                last_heading = text.strip()
                section_stack[lvl] = text.strip()
                # 清理更深层级
                for k in list(section_stack):
                    if k > lvl:
                        section_stack.pop(k, None)
                headings.append({"level": lvl, "text": text.strip(), "para_idx": pid})
                sections.append({"level": lvl, "text": text.strip(), "para_idx": pid})

            heading_by_para[para_idx] = last_heading
            sec = " / ".join(section_stack.get(i, "") for i in sorted(section_stack))

            # ---- fulltext（含题录行；表格另行标记）----
            if text.strip():
                fulltext_lines.append(f"{pid}{tag} {text.strip()}")

            if not text.strip():
                continue

            # ---- 题录 ----
            for m in CAPTION_RE.finditer(text):
                is_def = bool(CAPTION_DEF_RE.match(text)) or text.strip().startswith(m.group(0))
                captions.append({
                    "type": m.group(1), "num": m.group(2),
                    "text": m.group(3).strip()[:60],
                    "is_def": is_def, "para_idx": pid, "section": sec,
                    "raw": m.group(0),
                })

            # ---- 标准引用 ----
            book_spans = [(bm.group(1), bm.start(), bm.end()) for bm in GUIBOOK_RE.finditer(text)]
            for m in STD_CODE_RE.finditer(text):
                prefix, slashT, num, sep, year = m.group(1), m.group(2), m.group(3), m.group(4), m.group(5)
                raw = m.group(0)
                code = prefix + (slashT or "") + num + (("-" + year) if year else "")
                nc = norm_code(code)
                # 找最近《》名称
                best = None
                for name, bs, be in book_spans:
                    d = min(abs(m.start() - bs), abs(m.start() - be), abs(m.end() - bs), abs(m.end() - be))
                    if best is None or d < best[1]:
                        best = (name, d)
                name = best[0] if best else ""
                citations.append({
                    "kind": "std", "raw": raw.strip(),
                    "prefix": prefix, "slashT": (slashT or "").strip("/"),
                    "num": num, "year": year or "", "norm_code": nc,
                    "name": name, "norm_name": norm_name(name),
                    "para_idx": pid, "section": sec,
                    "context": _context(text, m.start(), m.end()),
                })

            # ---- 法规引用（《》名称 + 令号）----
            seen_book = set()
            for bm in GUIBOOK_RE.finditer(text):
                name = bm.group(1)
                key = (name, bm.start())
                if key in seen_book:
                    continue
                seen_book.add(key)
                # 判断是否法规/规章名称
                if re.search(r"(法|条例|规定|办法|细则|决定|意见|纲要|大纲|目录|清单)$", name) or \
                   re.search(r"(法|条例|规定|办法|管理办法|实施办法)$", name):
                    kind_law = "law"
                elif re.search(r"(规范|标准|规程|导则|通则|术语)$", name):
                    kind_law = "std_name_only"
                else:
                    kind_law = "other"
                citations.append({
                    "kind": kind_law, "raw": f"《{name}》", "norm_code": "",
                    "name": name, "norm_name": norm_name(name),
                    "para_idx": pid, "section": sec,
                    "context": _context(text, bm.start(), bm.end()),
                })
            for m in LAW_CODE_RE.finditer(text):
                citations.append({
                    "kind": "law_code", "raw": m.group(0).strip(), "norm_code": "",
                    "name": "", "norm_name": "",
                    "para_idx": pid, "section": sec,
                    "context": _context(text, m.start(), m.end()),
                })

            # ---- 数值 ----
            for rgx, key, grp in (
                (FLOW_RE, "flows", 1), (MODULUS_RE, "moduli", 1),
                (AREA_KM2_RE, "areas_km2", 1), (AREA_MU_RE, "areas_mu", 1),
                (BIGUNIT_RE, "bigunits", 1),
            ):
                for m in rgx.finditer(text):
                    val = m.group(grp)
                    numbers.append({
                        "key": key, "value": val,
                        "raw": m.group(0).strip(),
                        "para_idx": pid, "section": sec,
                        "context": _context(text, m.start(), m.end()),
                    })
            for m in POW10_RE.finditer(text):
                exp = m.group(1) or m.group(2) or ""
                numbers.append({
                    "key": "pow10", "value": exp,
                    "raw": m.group(0).strip(),
                    "para_idx": pid, "section": sec,
                    "context": _context(text, m.start(), m.end()),
                })

        elif kind == "table":
            tbl_idx += 1
            grid = grid_from_table(obj)
            nrows = len(grid)
            ncols = max((len(r) for r in grid), default=0)
            header = [ (str(c)[:40] if c is not None else "") for c in (grid[0] if grid else [])]
            # 邻近题录：回看最近 6 个段落里的"表 X.X-X"定义
            cap = ""
            cap_num = ""
            for cap_item in reversed(captions):
                if cap_item["type"] == "表" and cap_item["is_def"]:
                    cap = cap_item["text"]
                    cap_num = cap_item["num"]
                    break
            pid = f"T{tbl_idx:03d}"
            sec = " / ".join(section_stack.get(i, "") for i in sorted(section_stack))
            tables.append({
                "id": tbl_idx, "pid": pid, "caption": cap, "caption_num": cap_num,
                "nrows": nrows, "ncols": ncols, "header": header,
                "grid": grid, "para_idx_before": f"P{para_idx:04d}", "section": sec,
            })
            fulltext_lines.append(
                f"[P{para_idx+1:04d}] 【表{tbl_idx:03d}】({nrows}行×{ncols}列) "
                f"{cap}（详见 tables.json#{tbl_idx}）"
            )

    # ---- 写出 ----
    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, "fulltext.md"), "w", encoding="utf-8") as f:
        f.write("# 全文（段落序号 + 标题层级；表格仅留标记）\n\n")
        f.write("\n".join(fulltext_lines))
        f.write("\n")

    with open(os.path.join(out_dir, "tables.json"), "w", encoding="utf-8") as f:
        json.dump(tables, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out_dir, "captions.json"), "w", encoding="utf-8") as f:
        json.dump(captions, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out_dir, "citations.json"), "w", encoding="utf-8") as f:
        json.dump(citations, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out_dir, "numbers.json"), "w", encoding="utf-8") as f:
        json.dump(numbers, f, ensure_ascii=False, indent=1)

    meta = {
        "file": os.path.basename(docx_path),
        "docx_path": docx_path,
        "paragraphs": para_idx + 1,
        "tables": tbl_idx + 1,
        "headings": headings,
        "sections": sections,
    }
    with open(os.path.join(out_dir, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)

    # 定位映射：每段最近标题 + 每表题录号（页码由 page_map.py 补充）
    locators = {
        "heading_by_para": {str(k): v for k, v in heading_by_para.items()},
        "tables": {str(t["id"]): {"caption_num": t["caption_num"],
                                  "caption": t["caption"],
                                  "para_idx_before": t["para_idx_before"]}
                   for t in tables},
        "page_by_para": {},   # 由 page_map.py 填充（需 Word）
        "total_pages": None,
    }
    with open(os.path.join(out_dir, "locators.json"), "w", encoding="utf-8") as f:
        json.dump(locators, f, ensure_ascii=False, indent=1)

    return meta


def main():
    args = sys.argv[1:]
    if not args:
        print("用法: python extract_docx.py <docx路径> [输出目录]")
        sys.exit(1)
    docx_path = args[0]
    here = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.normpath(os.path.join(here, ".."))
    default_out = os.path.join(skill_root, "_run",
                               os.path.splitext(os.path.basename(docx_path))[0])
    out_dir = args[1] if len(args) > 1 else default_out
    meta = extract(docx_path, out_dir)
    print(f"[extract_docx] {meta['file']}: 段落 {meta['paragraphs']}，表格 {meta['tables']}，标题 {len(meta['headings'])}")
    print(f"[extract_docx] 输出目录 {out_dir}")


if __name__ == "__main__":
    main()
