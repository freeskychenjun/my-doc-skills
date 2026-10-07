# -*- coding: utf-8 -*-
"""
page_map.py — 用 Word COM 计算每个"顶层段落"对应的 Word 页码，写入 locators.json 的 page_by_para。

python-docx 不能从 XML 得到渲染页码（需 Word 分页）。本脚本在 Word 中打开文档，
按文档顺序遍历段落、跳过表格内段落（与抽取器 P 编号一致），记录每段页码。

依赖：Windows + Word（pywin32）。无 Word 时静默跳过，定位回退到章节标题。
用法：python page_map.py <run_dir>
"""

from __future__ import annotations
import json
import os
import sys

# GBK 控制台自愈：打印 ✓/✗ 等会 UnicodeEncodeError 且产物不落盘（2026-09-30 实测）
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass
import time


def build(docx_path, run_dir):
    import win32com.client as wc  # type: ignore
    WD_PAGE = 3        # wdActiveEndPageNumber
    WD_INTABLE = 12    # wdWithInTable
    word = wc.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        word.ScreenUpdating = False
    except Exception:
        pass
    t0 = time.time()
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        total = doc.ComputeStatistics(2)  # wdStatisticPages，触发分页
    except Exception:
        total = None
    page_by_para = {}
    i = -1
    for p in doc.Paragraphs:
        try:
            if p.Range.Information(WD_INTABLE):
                continue  # 跳过表格内段落，保持与抽取器顶层段落一致
        except Exception:
            continue
        i += 1
        try:
            page_by_para[i] = int(p.Range.Information(WD_PAGE))
        except Exception:
            page_by_para[i] = None
    try:
        doc.Close(False)
    except Exception:
        pass
    try:
        word.Quit()
    except Exception:
        pass
    print(f"[page_map] 打开+分页+遍历 {i + 1} 段，用时 {time.time() - t0:.1f}s，总页数 {total}")
    return page_by_para, total


def main():
    args = sys.argv[1:]
    run_dir = args[0] if args else None
    if not run_dir:
        print("用法: python page_map.py <run_dir>"); sys.exit(1)
    meta = json.load(open(os.path.join(run_dir, "meta.json"), encoding="utf-8"))
    docx_path = meta["docx_path"]
    loc_path = os.path.join(run_dir, "locators.json")
    locators = json.load(open(loc_path, encoding="utf-8")) if os.path.exists(loc_path) else {}
    try:
        page_by_para, total = build(docx_path, run_dir)
    except Exception as e:
        print(f"[page_map] Word COM 不可用或失败，跳过页码: {e!r}")
        return
    n_para = meta.get("paragraphs", 0)
    if i_count := (max(page_by_para) + 1 if page_by_para else 0) != n_para:
        print(f"[page_map] 警告：Word 顶层段落数({max(page_by_para)+1 if page_by_para else 0}) "
              f"与抽取段落数({n_para})不一致，缺失段落将以章节标题定位")
    locators["page_by_para"] = {str(k): v for k, v in page_by_para.items()}
    locators["total_pages"] = total
    json.dump(locators, open(loc_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"[page_map] 已写入 {loc_path}（{len(page_by_para)} 段落页码）")


if __name__ == "__main__":
    main()
