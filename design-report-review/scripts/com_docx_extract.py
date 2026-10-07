# -*- coding: utf-8 -*-
"""
com_docx_extract.py — 报告智能校审技能 · Word COM 内存直读器

当 .doc 无法通过 SaveAs 转 .docx 时（如终端安全软件 Agile DGS 拦截 Word 写盘，
导致输出文件被加密），用本脚本直接通过 Word COM 遍历文档对象模型，把段落和表格
导出为与 extract_docx.py 相同的 artifacts 格式，绕过文件写入。

输出（与 extract_docx.py 一致，供下游 deterministic_checks.py 使用）：
  fulltext.md     带段落序号与标题层级的全文
  tables.json     全部表格（2D 网格 + 邻近题录 + 所在章节）
  captions.json   图/表题录
  citations.json  标准与法规引用（由正则提取）
  numbers.json    关键数值（由正则提取）
  meta.json       文档元信息
  locators.json   定位映射

用法：
  python com_docx_extract.py "<doc或docx路径>" ["<输出目录>"]
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
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib_docx import (  # noqa: E402
    STD_CODE_RE, LAW_CODE_RE, GUIBOOK_RE, CAPTION_RE, CAPTION_DEF_RE,
    FLOW_RE, MODULUS_RE, AREA_KM2_RE, AREA_MU_RE, BIGUNIT_RE, POW10_RE,
    norm_code, norm_name,
)


HEADING_RE = re.compile(r"(?:Heading|标题|heading|Title)\s*([1-9])", re.IGNORECASE)


def _context(text, start, end, win=35):
    a = max(0, start - win)
    b = min(len(text), end + win)
    return text[a:b].replace("\n", " ").strip()


def _table_grid(table):
    """从 Word COM Table 对象读出二维网格 [[str,...],...]。

    用 table.Range.Cells 按单元格遍历（而非 Rows），可正确处理含合并单元格的
    不规则表格——每个单元格通过 cell.RowIndex / cell.ColumnIndex 定位。
    注意：合并单元格在 Word COM 里会重复返回（同一合并区在多个 (Row,Col) 位置
    都返回同一 cell 对象），用 id 去重，每个逻辑单元格只取一次。
    """
    try:
        cells = table.Range.Cells
    except Exception:
        return []
    # 用 (row, col) -> text 建网格，合并单元格取首次出现位置
    cell_map = {}
    seen_ids = set()
    max_row = max_col = 0
    try:
        for cell in cells:
            cid = id(cell._oleobj_) if hasattr(cell, '_oleobj_') else id(cell)
            try:
                ri = cell.RowIndex
                ci = cell.ColumnIndex
            except Exception:
                continue
            if (ri, ci) in cell_map:
                continue
            txt = cell.Range.Text or ""
            txt = txt.replace("\r", "").replace("\x07", "").strip()
            cell_map[(ri, ci)] = txt
            if ri > max_row:
                max_row = ri
            if ci > max_col:
                max_col = ci
    except Exception:
        pass
    # 展平为二维网格（1-based 转 0-based）
    grid = []
    for ri in range(1, max_row + 1):
        row = []
        for ci in range(1, max_col + 1):
            row.append(cell_map.get((ri, ci), ""))
        grid.append(row)
    return grid


def com_extract(doc_path: str, out_dir: str):
    import win32com.client  # type: ignore

    # 确保无残留 Word 进程
    os.system('taskkill /F /IM WINWORD.EXE >nul 2>&1')
    time.sleep(1)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    fulltext_lines = []
    tables, captions, citations, numbers = [], [], [], []
    headings, sections = [], []
    section_stack = {}      # level -> text
    heading_by_para = {}
    last_heading = ""

    try:
        d = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True, ConfirmConversions=False)
        n_paras = d.Paragraphs.Count
        n_tables = d.Tables.Count

        # 遍历文档主体内容（按段落顺序，表格穿插）
        # 用 Range 的 Paragraphs 与 Tables 分别遍历，再按位置合并
        # 更可靠的方式：遍历 d.Range.InlineShapes + 段落，但 Word COM 的
        # "段落与表格交错顺序"无法直接拿到。这里用文献常用方法：
        # 遍历 d.Range.Paragraphs，遇到属于某表格首段的就插入表格标记。

        # 先建表 -> 首段范围 映射（用表格第一行第一格的段落定位）
        tbl_start = {}
        for i in range(1, n_tables + 1):
            try:
                tbl = d.Tables(i)
                tbl_start[int(tbl.Range.Start)] = i
            except Exception:
                pass

        para_idx = -1
        tbl_idx = 0
        body = d.Content
        # 按字符位置顺序遍历段落
        all_paras = body.Paragraphs
        # 为提速，先取表格 Range 列表
        tbl_ranges = []
        for i in range(1, n_tables + 1):
            try:
                t = d.Tables(i)
                tbl_ranges.append((int(t.Range.Start), int(t.Range.End), i, t))
            except Exception:
                pass
        tbl_cursor = 0  # 下一个待插入的表格索引

        for p in all_paras:
            para_idx += 1
            try:
                pstart = int(p.Range.Start)
            except Exception:
                pstart = -1

            # 检查是否到了表格位置（表格 Start 之前的段落已处理完）
            while (tbl_cursor < len(tbl_ranges) and
                   pstart >= 0 and pstart >= tbl_ranges[tbl_cursor][0]):
                ts, te, tid, tbl = tbl_ranges[tbl_cursor]
                tbl_cursor += 1
                tbl_idx += 1
                if pstart == ts:
                    # 当前段落就是表格首段，跳过本段文本（用表格标记代替）
                    pass
                # 读取表格
                grid = _table_grid(tbl)
                nrows = len(grid)
                ncols = max((len(r) for r in grid), default=0)
                header = [(str(c)[:40] if c else "") for c in (grid[0] if grid else [])]
                cap = ""
                cap_num = ""
                for cap_item in reversed(captions):
                    if cap_item["type"] == "表" and cap_item["is_def"]:
                        cap = cap_item["text"]
                        cap_num = cap_item["num"]
                        break
                pid = f"T{tbl_idx - 1:03d}"
                sec = " / ".join(section_stack.get(i, "") for i in sorted(section_stack))
                tables.append({
                    "id": tbl_idx - 1, "pid": pid, "caption": cap,
                    "caption_num": cap_num, "nrows": nrows, "ncols": ncols,
                    "header": header, "grid": grid,
                    "para_idx_before": f"P{para_idx:04d}", "section": sec,
                })
                fulltext_lines.append(
                    f"[P{para_idx:04d}] 【表{tbl_idx - 1:03d}】({nrows}行×{ncols}列) "
                    f"{cap}（详见 tables.json#{tbl_idx - 1}）"
                )
                # 如果表格正好在当前段落位置，跳过当前段落（已是表格标记）
                if pstart == ts:
                    break
            else:
                # 处理普通段落
                text = ""
                try:
                    text = p.Range.Text or ""
                except Exception:
                    text = ""
                # 清理 Word 段落末尾的控制符
                text = text.replace("\r", "").replace("\x07", "").replace("\x0b", "").strip()

                # 标题层级（OutlineLevel 更可靠：1-9 为大纲级别，10=正文）
                lvl = None
                try:
                    ol = p.OutlineLevel
                    if ol and 1 <= ol <= 9:
                        lvl = ol
                except Exception:
                    pass
                # 退回用样式名
                if lvl is None:
                    try:
                        sname = p.Style.NameLocal or ""
                        m = HEADING_RE.search(sname)
                        if m:
                            lvl = int(m.group(1))
                        elif sname in ("标题", "Title"):
                            lvl = 1
                    except Exception:
                        pass

                tag = f"[H{lvl}]" if lvl else ""
                pid = f"P{para_idx:04d}"

                if lvl and text:
                    last_heading = text
                    section_stack[lvl] = text
                    for k in list(section_stack):
                        if k > lvl:
                            section_stack.pop(k, None)
                    headings.append({"level": lvl, "text": text, "para_idx": pid})
                    sections.append({"level": lvl, "text": text, "para_idx": pid})

                heading_by_para[para_idx] = last_heading
                sec = " / ".join(section_stack.get(i, "") for i in sorted(section_stack))

                if text:
                    fulltext_lines.append(f"{pid}{tag} {text}")
                    _extract_inline(text, pid, sec, captions, citations, numbers)
                continue

        d.Close(False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        os.system('taskkill /F /IM WINWORD.EXE >nul 2>&1')

    # ---- 写出（与 extract_docx.py 一致）----
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
        "file": os.path.basename(doc_path),
        "docx_path": doc_path,
        "paragraphs": para_idx + 1,
        "tables": tbl_idx,
        "headings": headings,
        "sections": sections,
    }
    with open(os.path.join(out_dir, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)

    locators = {
        "heading_by_para": {str(k): v for k, v in heading_by_para.items()},
        "tables": {str(t["id"]): {"caption_num": t["caption_num"],
                                  "caption": t["caption"],
                                  "para_idx_before": t["para_idx_before"]}
                   for t in tables},
        "page_by_para": {},
        "total_pages": None,
    }
    with open(os.path.join(out_dir, "locators.json"), "w", encoding="utf-8") as f:
        json.dump(locators, f, ensure_ascii=False, indent=1)

    return meta


def _extract_inline(text, pid, sec, captions, citations, numbers):
    """与 extract_docx.py 相同的内联提取逻辑（题录/引用/数值）。"""
    # 题录
    for m in CAPTION_RE.finditer(text):
        is_def = bool(CAPTION_DEF_RE.match(text)) or text.strip().startswith(m.group(0))
        captions.append({
            "type": m.group(1), "num": m.group(2),
            "text": m.group(3).strip()[:60],
            "is_def": is_def, "para_idx": pid, "section": sec, "raw": m.group(0),
        })
    # 标准引用
    book_spans = [(bm.group(1), bm.start(), bm.end()) for bm in GUIBOOK_RE.finditer(text)]
    for m in STD_CODE_RE.finditer(text):
        prefix, slashT, num, year = m.group(1), m.group(2), m.group(3), m.group(5)
        raw = m.group(0)
        code = prefix + (slashT or "") + num + (("-" + year) if year else "")
        nc = norm_code(code)
        best = None
        for name, bs, be in book_spans:
            d = min(abs(m.start() - bs), abs(m.start() - be), abs(m.end() - bs), abs(m.end() - be))
            if best is None or d < best[1]:
                best = (name, d)
        name = best[0] if best else ""
        citations.append({
            "kind": "std", "raw": raw.strip(), "prefix": prefix,
            "slashT": (slashT or "").strip("/"), "num": num, "year": year or "",
            "norm_code": nc, "name": name, "norm_name": norm_name(name),
            "para_idx": pid, "section": sec,
            "context": _context(text, m.start(), m.end()),
        })
    # 法规引用
    seen_book = set()
    for bm in GUIBOOK_RE.finditer(text):
        name = bm.group(1)
        key = (name, bm.start())
        if key in seen_book:
            continue
        seen_book.add(key)
        if re.search(r"(法|条例|规定|办法|细则|决定|意见|纲要|大纲|目录|清单)$", name):
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
    # 数值
    for rgx, key, grp in (
        (FLOW_RE, "flows", 1), (MODULUS_RE, "moduli", 1),
        (AREA_KM2_RE, "areas_km2", 1), (AREA_MU_RE, "areas_mu", 1),
        (BIGUNIT_RE, "bigunits", 1),
    ):
        for m in rgx.finditer(text):
            numbers.append({
                "key": key, "value": m.group(grp), "raw": m.group(0).strip(),
                "para_idx": pid, "section": sec,
                "context": _context(text, m.start(), m.end()),
            })
    for m in POW10_RE.finditer(text):
        exp = m.group(1) or m.group(2) or ""
        numbers.append({
            "key": "pow10", "value": exp, "raw": m.group(0).strip(),
            "para_idx": pid, "section": sec,
            "context": _context(text, m.start(), m.end()),
        })


def main():
    args = sys.argv[1:]
    if not args:
        print("用法: python com_docx_extract.py <doc路径> [输出目录]")
        sys.exit(1)
    doc_path = args[0]
    here = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.normpath(os.path.join(here, ".."))
    # 输出目录名去掉末尾空格/点（Windows 不允许）
    base = os.path.splitext(os.path.basename(doc_path))[0].rstrip(" .")
    default_out = os.path.join(skill_root, "_run", base)
    out_dir = args[1] if len(args) > 1 else default_out
    out_dir = out_dir.rstrip(" .")
    meta = com_extract(doc_path, out_dir)
    print(f"[com_docx_extract] {meta['file']}: 段落 {meta['paragraphs']}，"
          f"表格 {meta['tables']}，标题 {len(meta['headings'])}")
    print(f"[com_docx_extract] 输出目录 {out_dir}")


if __name__ == "__main__":
    main()
