# -*- coding: utf-8 -*-
"""
parse_requirements.py — 解析 检查要求.md 为结构化 requirements.json。

输出 list[ {category, subcategory, id, name, content} ]，供编排器逐条分发与报告引用原文。
用法：python parse_requirements.py [检查要求.md] [输出json]
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


def parse(path: str):
    reqs = []
    category = ""
    subcategory = ""
    with open(path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.rstrip("\n")
            s = line.strip()
            if s.startswith("# ") and not s.startswith("# 检查结果输出要求"):
                category = s.lstrip("# ").strip()
                subcategory = ""
                continue
            if s.startswith("## "):
                subcategory = s.lstrip("# ").strip()
                continue
            if not s.startswith("|"):
                continue
            if re.fullmatch(r"[\s\-:|]+", s.strip("|")):
                continue
            cells = [c.strip() for c in s.strip("|").split("|")]
            # 表头行：仅当首列就是"编号"时跳过（避免误吞名称含"编号"的数据行）
            if cells and cells[0] == "编号":
                continue
            if len(cells) < 3:
                continue
            rid, name, content = cells[0], cells[1], cells[2]
            # 编号形如 mp_001 / ci_flood_001 / ds_001（字母段+下划线，末段为数字）
            if not (re.match(r"^[A-Za-z]+(?:_[A-Za-z0-9]+)+$", rid) and rid.split("_")[-1].isdigit()):
                continue
            reqs.append({
                "category": category, "subcategory": subcategory,
                "id": rid, "name": name, "content": content,
            })
    return reqs


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.normpath(os.path.join(here, ".."))
    args = sys.argv[1:]
    src = args[0] if len(args) > 0 else os.path.join(skill_root, "references", "检查要求.md")
    out = args[1] if len(args) > 1 else os.path.join(here, "requirements.json")
    reqs = parse(src)
    # 按 category 分组统计
    from collections import Counter, OrderedDict
    by_cat = OrderedDict()
    for r in reqs:
        by_cat.setdefault(r["category"], []).append(r["id"])
    json.dump(reqs, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"[parse_requirements] 共解析 {len(reqs)} 条检查要求 -> {out}")
    for cat, ids in by_cat.items():
        print(f"  {cat}: {len(ids)} 条 ({ids[0]}..{ids[-1]})")


if __name__ == "__main__":
    main()
