# -*- coding: utf-8 -*-
"""
slice_artifacts.py — 报告智能校审技能 · 按分组裁剪产物（提速核心）

把抽取产物 + 确定性证据按 4 个 subagent 分组裁剪成小文件，避免每个 agent 都
灌入 fulltext.md（最大 305KB）/ numbers.json（317KB）/ references_index.json（460KB）
等大文件。每个 slice 控制在 <30KB（约 8000 token）。

分组（与 SKILL.md 第 3 节一致）：
  groupA1 强条 mp_001–011（11 条）
  groupA2 常见 ci_*（9 条）
  groupB  一致性 cons_001–003 + 表格 tbl_001–003
  groupC  语法 gram_001–006/008 + 文字 std_002
  groupD  标准 ds_001–002 + 法规 lr_001–002
（A1/A2 共用同一份 A 数据切片，与拆分前 groupA 完全一致，仅要求清单不同——
  拆组只为把原 20 条的最重组负载减半，墙钟≈各组最大值，不丢任何证据）
（std_001/003/004、gram_007 由 det_findings.py 直出，不派 subagent，故不进 slice）

每个切片内嵌本组 "requirements" 字段（id/name/content 原文，取自
requirements.json）——下发提示词直接引用切片内清单，杜绝手抄漏条
（2026-09-30 实战曾手抄漏发 ci_runoff_002/ci_water_002/ci_water_003 共 3 条）。
生成前做分组覆盖自检：5 组清单 ∪ det 直出 4 条必须恰好等于全部要求，
不等即分组定义与 requirements.json 漂移，立即报错退出。

用法：
  python slice_artifacts.py <run_dir>
  → <run_dir>/_slices/groupA1.json / groupA2.json / groupB.json / groupC.json / groupD.json
"""

from __future__ import annotations
import json
import os
import re
import sys

# 直出条目（det_findings.py 直出，不进任何切片）
DET_DIRECT_IDS = {"std_001", "std_003", "std_004", "gram_007"}

# 分组 → 要求 id 过滤规则（分组定义与 SKILL.md 第 3 节一致；
# 切片内 requirements 清单由本表从 requirements.json 程序化生成，禁止手抄）
GROUP_ID_RULES = {
    "A1": lambda rid: rid.startswith("mp_"),
    "A2": lambda rid: rid.startswith("ci_"),
    "B":  lambda rid: rid.startswith(("cons_", "tbl_")),
    "C":  lambda rid: (rid.startswith("gram_") and rid not in DET_DIRECT_IDS) or rid == "std_002",
    "D":  lambda rid: rid.startswith(("ds_", "lr_")),
}


def _console_safe():
    # GBK 控制台下打印 ✓/✗ 会 UnicodeEncodeError 且中断写入，errors=replace 自愈
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(errors="replace")
        except Exception:
            pass


def _group_requirements(here):
    """按 GROUP_ID_RULES 从 requirements.json 生成各组要求清单，并做覆盖自检。"""
    reqs = _read_json(os.path.join(here, "requirements.json")) or []
    by_group = {g: [{"id": r["id"], "name": r.get("name", ""), "content": r.get("content", "")}
                    for r in reqs if "id" in r and rule(r["id"])]
                for g, rule in GROUP_ID_RULES.items()}
    covered = {r["id"] for lst in by_group.values() for r in lst} | DET_DIRECT_IDS
    all_ids = {r["id"] for r in reqs if "id" in r}
    if covered != all_ids:
        miss, extra = sorted(all_ids - covered), sorted(covered - all_ids)
        print(f"[slice] ✗ 分组覆盖自检失败：requirements 有而分组未覆盖 {miss}；"
              f"分组有而 requirements 无 {extra}。分组定义与要求清单漂移，禁止继续——"
              f"这正是 2026-09-30 漏发 3 条 ci_* 的根因，请核对 GROUP_ID_RULES。")
        sys.exit(1)
    return by_group


# ---- groupA 关键词：强条+常见设计问题涉及的水文主题 ----
# 用更具体的短语，避免"资料/合理性/系列"等在水利报告里高频的泛词把全文都选中
GROUPA_KW = re.compile(
    r"(设计洪水|暴雨|潮位|径流|参证站|水文比拟|水位流量|水面线|壅水|排涝|模数|"
    r"产流|汇流|洪痕|洪峰|插补|延长|考证|分期|率定|汇流参数|"
    r"三性|可靠性|一致性分析|合理性检查|合理性分析)"
)

# 各 slice 的硬上限（条数），超过即截断，防止大报告撑爆上下文
MAX_FULLTEXT_SLICE_LINES = 400      # groupA 关键词切片最多留这么多行
MAX_NUMBERS_FLOW_MOD = 80           # groupA flows/moduli 数值最多留这么多条
MAX_TABLES_IN_SLICE = 40            # groupB 表格最多留这么多张
MAX_NUMERIC_LINES = 120             # groupB 正文数字行最多留这么多行
MAX_CITATIONS_IN_SLICE = 120        # groupD citations 全量上限
MAX_FULLTEXT_GRAM_CHARS = 60000     # groupC 全文超过这么多字符则改为"标题+疑点采样+Grep指示"
# groupC 大报告时的疑点关键词：聚焦"可能有问题"的特征，不含高频通用标点（。；，：！？、）
# 否则中文正文每行都命中，等于没裁剪。
GRAM_SUSPECT_KW = re.compile(
    r"(图\s*\d|表\s*\d|《|》|第\s*[一二三四五六七八九十百千]+\s*[章条款节编号]|"
    r"等别|级别|m³|m3|km²|km2|10\s*的|Ⅲ|Ⅳ|Ⅱ|Ⅰ|Ⅴ|VI|IV|III|II)"
)

# 表格 grid 太大时按行裁剪（保留首行表头 + 前 N 行数据）
MAX_GRID_ROWS = 20
# det_table info 项每表最多保留条数（够 LLM 抽样判断即可）
MAX_INFO_PER_TABLE = 3
# references_index 命中项每类最多保留条数
MAX_IDX_HITS = 30


def _read_json(path):
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _read_text(path):
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        return f.read()


def _trim_grid(grid, max_rows=MAX_GRID_ROWS):
    """表格 grid 行数过多时只保留表头 + 前 max_rows-1 行，附截断标记。"""
    if not grid or len(grid) <= max_rows:
        return grid
    out = grid[:max_rows]
    # 末行加截断提示（不破坏列对齐：填 None 到 ncols）
    ncols = max(len(r) for r in grid)
    trunc = [None] * ncols
    trunc[0] = f"…（共 {len(grid)} 行，已截断显示前 {max_rows - 1} 行数据）"
    out.append(trunc)
    return out


# ============================================================
# groupA 数据切片：强条 mp + 常见 ci 共用（A1/A2 各写一份，内容相同）
# 需要全文里与水文主题相关的段落 + 流量/模数数值 + std003 硬违规
# ============================================================
def build_groupA(run_dir):
    fulltext = _read_text(os.path.join(run_dir, "fulltext.md"))
    numbers = _read_json(os.path.join(run_dir, "numbers.json")) or []
    det_numbers = _read_json(os.path.join(run_dir, "det_numbers.json")) or {}

    # fulltext 切片：标题行（[H*]）必留，正文行按关键词留，限量截断
    kept_lines = []
    n_total, n_kept = 0, 0
    for line in fulltext.splitlines():
        s = line.strip()
        if not s:
            continue
        n_total += 1
        is_heading = "[H" in line[:12]   # 形如 P0012[H3]
        # 表格标记行【表###】也留，便于定位
        is_table_marker = "【表" in line
        if is_heading or is_table_marker or GROUPA_KW.search(s):
            kept_lines.append(line)
            n_kept += 1
            if len(kept_lines) >= MAX_FULLTEXT_SLICE_LINES:
                break
    fulltext_slice = "\n".join(kept_lines)

    # numbers 只留 flows / moduli（排涝模数是 ci_drainage_001 的关键），限量
    nums_all = [n for n in numbers if n.get("key") in ("flows", "moduli")]
    # 优先留 moduli（ci_drainage_001 关键），其次 flows
    moduli = [n for n in nums_all if n.get("key") == "moduli"]
    flows = [n for n in nums_all if n.get("key") == "flows"]
    nums_slice = (moduli + flows)[:MAX_NUMBERS_FLOW_MOD]

    # std003 硬违规清单（供 ci 组参考流量精度，非判定 ci 条目）
    std003 = det_numbers.get("std003_flow", {})
    std003_brief = {
        "n_violations": std003.get("n_violations", 0),
        "violations": std003.get("violations", [])[:15],
    }

    return {
        "_doc": "强条 mp_001–011 + 常见设计问题 ci_* 所需切片",
        "_fulltext_stats": {"total_lines": n_total, "kept_lines": n_kept},
        "fulltext_slice": fulltext_slice,
        "numbers_flows_moduli": nums_slice,
        "det_std003_brief": std003_brief,
    }


# ============================================================
# groupB：一致性 cons + 表格 tbl
# 需要 tables 的 grid + det_table 的 fail/warn/info 摘要
# ============================================================
def build_groupB(run_dir):
    tables = _read_json(os.path.join(run_dir, "tables.json")) or []
    det_table = _read_json(os.path.join(run_dir, "det_table.json")) or []
    fulltext = _read_text(os.path.join(run_dir, "fulltext.md"))

    # tables：保留 id/caption/grid。
    # 表多时优先留"有 det 问题的 + 小表"，避免 98 张表全量撑爆 slice
    det_by_id = {t.get("id"): t for t in det_table}
    flagged_ids = {t.get("id") for t in det_table
                   if t.get("n_fail", 0) > 0 or t.get("n_warn", 0) > 0}

    def _table_priority(tb):
        # 有问题的表优先；其次小表（行数少）
        tid = tb.get("id")
        return (0 if tid in flagged_ids else 1, tb.get("nrows", 0))

    tables_sorted = sorted(tables, key=_table_priority)
    tables_slice = []
    for tb in tables_sorted:
        if len(tables_slice) >= MAX_TABLES_IN_SLICE:
            break
        tid = tb.get("id")
        is_flagged = tid in flagged_ids
        entry = {
            "id": tid,
            "pid": tb.get("pid"),
            "caption": tb.get("caption", ""),
            "caption_num": tb.get("caption_num", ""),
            "nrows": tb.get("nrows"),
            "ncols": tb.get("ncols"),
        }
        # 有问题的表给全 grid（截断）；没问题的表不给 grid（det 已判 pass，省空间）
        if is_flagged:
            entry["grid"] = _trim_grid(tb.get("grid", []))
        else:
            entry["grid_omitted"] = True   # 提示 LLM：如需看此表 grid 用 tables.json
        tables_slice.append(entry)
    if len(tables) > len(tables_slice):
        tables_slice.append({"_note": f"原文档共 {len(tables)} 张表，"
                                       f"此处只给优先级最高的 {len(tables_slice)} 张"
                                       f"（含全部 {len(flagged_ids)} 张有问题表）"})

    # det_table 摘要：每表给 n_checks/n_fail/n_warn + fail/warn 全列 + info 前 N 条
    det_brief = []
    for t in det_table:
        checks = t.get("checks", [])
        fails = [c for c in checks if c.get("status") == "fail"]
        warns = [c for c in checks if c.get("status") == "warn"]
        infos = [c for c in checks if c.get("status") == "info"][:MAX_INFO_PER_TABLE]
        det_brief.append({
            "id": t.get("id"),
            "pid": t.get("pid"),
            "caption": t.get("caption", ""),
            "n_checks": t.get("n_checks", 0),
            "n_fail": t.get("n_fail", 0),
            "n_warn": t.get("n_warn", 0),
            "fails": fails,
            "warns": warns,
            "infos_sample": infos,
        })

    # 一致性 cons_* 需要正文数据片段：保留含数字的正文行（便于文表对照），限量
    cons_lines = []
    for line in fulltext.splitlines():
        s = line.strip()
        if not s or "[H" in line[:12]:
            continue
        # 含阿拉伯数字的正文行（潜在数据陈述）
        if re.search(r"\d+(?:\.\d+)?", s):
            cons_lines.append(line)
            if len(cons_lines) >= MAX_NUMERIC_LINES:
                break

    return {
        "_doc": "一致性 cons_001–003 + 表格 tbl_001–003 所需切片",
        "tables": tables_slice,
        "det_table_brief": det_brief,
        "fulltext_numeric_lines": cons_lines,
    }


# ============================================================
# groupC：语法 gram + 文字 std_002
# 语法需通读，给全文（已剔除大表格，只有标记行）；附 det_numbering
# ============================================================
def build_groupC(run_dir):
    fulltext = _read_text(os.path.join(run_dir, "fulltext.md"))
    det_numbering = _read_json(os.path.join(run_dir, "det_numbering.json")) or {}
    det_numbers = _read_json(os.path.join(run_dir, "det_numbers.json")) or {}
    captions = _read_json(os.path.join(run_dir, "captions.json")) or []

    grade = det_numbers.get("std004_grade", [])

    # 语法/文字需要通读，但大报告全文会撑爆上下文。
    # 策略：全文 <= 60KB 时全量给；否则只给"标题结构 + 疑点行采样"，
    # 并明确指示 subagent 用 Grep 分段读 fulltext.md 检查标点/错字（按章节轮询）。
    if len(fulltext) <= MAX_FULLTEXT_GRAM_CHARS:
        fulltext_out = fulltext
        fulltext_note = ""
    else:
        heading_lines = []
        suspect_lines = []
        for line in fulltext.splitlines():
            s = line.strip()
            if not s:
                continue
            if "[H" in line[:12]:
                heading_lines.append(line)
            elif GRAM_SUSPECT_KW.search(s):
                suspect_lines.append(line)
        # 疑点行限量（大报告可能命中上百行）
        suspect_kept = suspect_lines[:200]
        fulltext_out = "\n".join(heading_lines + suspect_kept)
        fulltext_note = (
            f"全文 {len(fulltext)} 字符过大，无法一次灌入。此处提供：\n"
            f"  (1) 全部 {len(heading_lines)} 个标题行（章节结构，便于定位）；\n"
            f"  (2) {len(suspect_kept)} 个疑点行采样（含图表编号/《》/罗马数字/量纲等，"
            f"原始疑点共 {len(suspect_lines)} 行）。\n"
            f"gram_001–006（语法/标点/错字/通顺）检查方法：用 Grep 在 "
            f"<run_dir>/fulltext.md 按段落号区间分段（如 P0000–P0500、P0500–P1000）"
            f"逐段读取并检查，每段报具体 loc+original+suggestion。勿一次性读全文。"
        )

    return {
        "_doc": "语法 gram_001–006/008 + 文字 std_002 所需切片",
        "_fulltext_note": fulltext_note,
        "fulltext": fulltext_out,
        "captions": [{"type": c.get("type"), "num": c.get("num"),
                      "text": c.get("text"), "para_idx": c.get("para_idx")}
                     for c in captions],
        "det_numbering": det_numbering,
        "det_std004_grade": grade,
    }


# ============================================================
# groupD：标准 ds + 法规 lr
# citations 全量 + det_citations 摘要 + references_index 命中项
# ============================================================
def build_groupD(run_dir, here):
    citations = _read_json(os.path.join(run_dir, "citations.json")) or []
    det_cit = _read_json(os.path.join(run_dir, "det_citations.json")) or {}
    idx_path = os.path.join(here, "references_index.json")
    idx = _read_json(idx_path) or {}

    # det_citations 摘要：summary + consistency + 非 valid 的 citations 明细
    cit_list = det_cit.get("citations", [])
    suspicious = [c for c in cit_list
                  if c.get("validity") not in ("valid", "valid_name", None)]
    det_brief = {
        "summary": det_cit.get("summary", {}),
        "consistency": det_cit.get("consistency", []),
        "suspicious_citations": suspicious,   # not_in_list/year_mismatch/name_mismatch
    }

    # citations：可疑的全留，valid 的只采样（每类各取若干条去重后代表样）
    suspicious_raws = {c.get("raw") for c in citations
                       if any(d.get("raw") == c.get("raw") for d in suspicious)}
    valid_sample = []
    seen_norm = set()
    for c in citations:
        if c.get("raw") in suspicious_raws:
            continue
        nc = c.get("norm_code") or c.get("norm_name") or c.get("raw")
        if nc in seen_norm:
            continue
        seen_norm.add(nc)
        valid_sample.append(c)
        if len(valid_sample) >= 30:
            break
    citations_slice = [c for c in citations if c.get("raw") in suspicious_raws] + valid_sample
    if len(citations) > len(citations_slice):
        citations_slice.append({"_note": f"原文档共 {len(citations)} 条引用，"
                                         f"此处给全部可疑 + {len(valid_sample)} 条 valid 采样，"
                                         f"完整列表见 citations.json"})

    # references_index 只留与本文档 citations 命中的条目
    doc_codes = set()
    doc_names = set()
    doc_prefixes = set()
    for c in citations:
        nc = c.get("norm_code", "")
        nm = c.get("norm_name") or c.get("name", "")
        if nc:
            doc_codes.add(nc)
            doc_prefixes.add(re.sub(r"-\d{4}$", "", nc))
        if nm:
            doc_names.add(nm)

    by_code = idx.get("standards_by_code", {})
    by_name = idx.get("standards_by_name", {})
    by_prefix = idx.get("names_by_prefix_code", {})
    law_by_name = idx.get("laws_by_name", {})
    std_list = idx.get("standards", [])
    law_list = idx.get("laws", [])

    def _resolve(idxs, lst):
        """idx 列表（元素可能为 int 或 str 数字）→ 原始条目列表，越界跳过。"""
        out = []
        for i in idxs:
            try:
                k = int(i)
            except (TypeError, ValueError):
                continue
            if 0 <= k < len(lst):
                out.append(lst[k])
        return out

    hit_codes = {k: _resolve(by_code[k], std_list) for k in doc_codes if k in by_code}
    hit_names = {nm: _resolve(by_name[nm], std_list)
                 for nm in by_name if nm in doc_names}
    hit_prefix = {k: by_prefix[k][:3] for k in doc_prefixes if k in by_prefix}
    hit_law_names = {nm: _resolve(law_by_name[nm], law_list)
                     for nm in law_by_name if nm in doc_names}

    # 各类限量，防止单类爆炸
    def _cap(d, n=MAX_IDX_HITS):
        items = list(d.items())[:n]
        return dict(items)

    idx_hits = {
        "standards_by_code_hit": _cap(hit_codes),
        "standards_by_name_hit": _cap(hit_names),
        "standards_by_prefix_hit": _cap(hit_prefix),
        "laws_by_name_hit": _cap(hit_law_names),
        "_note": "只含本文档实际引用到的标准/法规；未命中的 norm_code 表示 not_in_list",
    }

    return {
        "_doc": "标准 ds_001–002 + 法规 lr_001–002 所需切片",
        "citations": citations_slice,
        "det_citations_brief": det_brief,
        "references_index_hits": idx_hits,
    }


def _write_slice(out_dir, name, data, n_reqs):
    p = os.path.join(out_dir, name)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    size_kb = os.path.getsize(p) / 1024
    print(f"[slice] {name}: {size_kb:.1f} KB（含 {n_reqs} 条要求清单）")


def main():
    _console_safe()
    args = sys.argv[1:]
    run_dir = args[0] if args else None
    if not run_dir or not os.path.isdir(run_dir):
        print("用法: python slice_artifacts.py <run_dir>")
        sys.exit(1)
    here = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(run_dir, "_slices")
    os.makedirs(out_dir, exist_ok=True)

    # 各组要求清单（程序化取自 requirements.json + 覆盖自检，漂移即退出）
    reqs_by_group = _group_requirements(here)

    for name, fn, gkey in (("groupB.json", build_groupB, "B"),
                           ("groupC.json", build_groupC, "C")):
        data = fn(run_dir)
        data["requirements"] = reqs_by_group[gkey]
        _write_slice(out_dir, name, data, len(reqs_by_group[gkey]))

    # A 组数据一份，A1（强条 mp）/A2（常见 ci）两个组共用：
    # 数据切片与拆分前 groupA 完全一致，仅 _doc 与 requirements 清单不同，拆组不减证据。
    data_a = build_groupA(run_dir)
    for name, gkey, doc in (("groupA1.json", "A1", "强条 mp_001–011（A1 组）所需切片"),
                            ("groupA2.json", "A2", "常见设计问题 ci_*（A2 组）所需切片")):
        data = dict(data_a)
        data["_doc"] = doc
        data["requirements"] = reqs_by_group[gkey]
        _write_slice(out_dir, name, data, len(reqs_by_group[gkey]))

    groupD = build_groupD(run_dir, here)
    groupD["requirements"] = reqs_by_group["D"]
    _write_slice(out_dir, "groupD.json", groupD, len(reqs_by_group["D"]))

    print(f"[slice_artifacts] 5 个切片已生成到 {out_dir}"
          f"（内嵌要求清单 {sum(len(v) for k, v in reqs_by_group.items())} 条 + det 直出 "
          f"{len(DET_DIRECT_IDS)} 条 = 全覆盖）")


if __name__ == "__main__":
    main()
