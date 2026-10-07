# -*- coding: utf-8 -*-
"""
build_reference_index.py — 解析"水利标准.md"与"法律法规.md"，构建标准化检索索引。

输出 references_index.json：
  {
    "standards": [ {序号, 编号, 名称, 实施日期, 备注, norm_code, norm_name}, ... ],
    "laws":      [ ... ],
    "standards_by_code": {norm_code: [序号, ...]},
    "standards_by_name": {norm_name: [序号, ...]},
    "laws_by_code":      {norm_code: [序号, ...]},
    "laws_by_name":      {norm_name: [序号, ...]},
    "names_by_prefix_code": { "SL252": [序号,...], ... }  # 仅按"前缀+号"查(忽略年份)
  }

用法：
  python build_reference_index.py 水利标准.md 法律法规.md [输出json]
  （缺省输出到脚本同目录的 references_index.json）
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
from lib_docx import norm_code, norm_name, STD_CODE_RE, LAW_CODE_RE  # noqa: E402


def parse_md_table(path: str):
    """解析 markdown 表格，返回 list[list[str]]（每个内层 list 为一行的各列去空白文本）。"""
    rows = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            s = line.rstrip("\n").strip()
            if not s.startswith("|"):
                continue
            # 跳过分隔行 | --- | --- |
            inner = s.strip("|")
            if re.fullmatch(r"[\s\-:|]+", inner):
                continue
            cells = [c.strip() for c in s.strip("|").split("|")]
            rows.append(cells)
    return rows


def _prefix_num(normcode: str) -> str:
    """从归一化编号取"前缀+号"（去掉年份），用于跨年份匹配。

    'SL/T252-2017' -> 'SL/T252'
    'GB50265-2022' -> 'GB50265'
    """
    return re.sub(r"-\d{4}$", "", normcode)


def build(std_path: str, law_path: str, out_path: str):
    standards, laws = [], []

    # ---- 水利标准：序号 | 标准编号 | 标准名称 | 实施日期 ----
    rows = parse_md_table(std_path)
    for cells in rows:
        if len(cells) < 4:
            continue
        序号, 编号, 名称, 实施日期 = cells[0], cells[1], cells[2], cells[3]
        if not re.search(r"\d", 编号) and not re.search(r"\d", 名称):
            continue
        standards.append({
            "序号": 序号,
            "编号": 编号,
            "名称": 名称,
            "实施日期": 实施日期,
            "备注": "",
            "norm_code": norm_code(编号),
            "norm_name": norm_name(名称),
        })

    # ---- 法律法规：序号 | 标准编号 | 标准名称 | 实施日期 | 备注 ----
    rows = parse_md_table(law_path)
    for cells in rows:
        if len(cells) < 4:
            continue
        序号 = cells[0]
        编号 = cells[1]
        名称 = cells[2]
        实施日期 = cells[3]
        备注 = cells[4] if len(cells) > 4 else ""
        if not 名称 or 名称 in ("标准名称",):
            continue
        if not re.search(r"[一-龥]", 名称):
            continue
        laws.append({
            "序号": 序号,
            "编号": 编号,
            "名称": 名称,
            "实施日期": 实施日期,
            "备注": 备注,
            "norm_code": norm_code(编号),
            "norm_name": norm_name(名称),
        })

    # ---- 建立索引 ----
    def idx_by(entries, keyfn):
        d = {}
        for i, e in enumerate(entries):
            k = keyfn(e)
            if not k:
                continue
            d.setdefault(k, []).append(e["序号"])
        return d

    standards_by_code = idx_by(standards, lambda e: e["norm_code"])
    standards_by_name = idx_by(standards, lambda e: e["norm_name"])
    laws_by_code = idx_by(laws, lambda e: e["norm_code"])
    laws_by_name = idx_by(laws, lambda e: e["norm_name"])

    # 前缀+号 -> 序号（用于发现"年份不一致/已更新"）
    names_by_prefix_code = {}
    for e in standards:
        pn = _prefix_num(e["norm_code"])
        if pn:
            names_by_prefix_code.setdefault(pn, []).append(
                {"序号": e["序号"], "编号": e["编号"], "名称": e["名称"],
                 "实施日期": e["实施日期"], "norm_code": e["norm_code"]})

    index = {
        "standards": standards,
        "laws": laws,
        "standards_by_code": standards_by_code,
        "standards_by_name": standards_by_name,
        "laws_by_code": laws_by_code,
        "laws_by_name": laws_by_name,
        "names_by_prefix_code": names_by_prefix_code,
        "stats": {
            "standards": len(standards),
            "laws": len(laws),
        },
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=1)
    return index


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.normpath(os.path.join(here, ".."))
    args = sys.argv[1:]
    std_path = args[0] if len(args) > 0 else os.path.join(skill_root, "references", "水利标准.md")
    law_path = args[1] if len(args) > 1 else os.path.join(skill_root, "references", "法律法规.md")
    out_path = args[2] if len(args) > 2 else os.path.join(skill_root, "scripts", "references_index.json")
    idx = build(std_path, law_path, out_path)
    print(f"[build_reference_index] 标准 {idx['stats']['standards']} 条，法规 {idx['stats']['laws']} 条")
    print(f"[build_reference_index] 写入 {out_path}")


if __name__ == "__main__":
    main()
