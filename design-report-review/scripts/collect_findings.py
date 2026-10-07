# -*- coding: utf-8 -*-
"""
collect_findings.py — 报告智能校审技能 · 收集+合并+校验（流程固化核心）

统一收集 5 个 subagent 分组的产物 + det_findings 的直出结论，合并成
findings.json，消除历史上"主 agent 现场拼 build_findings.py/_merge.py"的绕路。

输入：
  <run_dir>/groupA1/A2/B/C/D.json   subagent 各组输出（JSON，含 category+results）
        （默认在 <run_dir> 下查找；--from 可指定子目录，如 _slices/）
  <run_dir>/det_findings.json            确定性直出结论（优先级最高）
  $SKILL/scripts/requirements.json       42 条要求（用于覆盖率校验 + name 回填）

校验规则：
  - 每个 result 必备 id；缺失字段补默认（verdict→未检出，issues/evidence→[]）
  - verdict ∈ {符合,部分符合,不符合,不适用}，否则归一为"未检出"
  - issues 元素保留 loc/original/type/suggestion；evidence 元素保留 loc/quote
  - det_findings 的同 id 结果覆盖 subagent 结果（det 优先）

覆盖率（硬门禁，2026-10-01 起）：对照 requirements.json 全部 42 条；仍有缺失时
        findings.json 照写（缺失条目按"未检出"占位、_source=missing），但打印 FAIL
        并以退出码 1 结束——占位 ≠ 完成，须按 SKILL.md ③ 的失败恢复三梯度补齐后
        重跑本脚本，不得带占位进入合成。确属无法补判时加 --allow-missing 放行
        （最终报告须向用户说明）。背景：2026-09-30 实战曾因手抄下发清单漏发 3 条
        ci_*，靠人工汇总校验才兜住——静默占位会掩盖这类丢失。

用法：
  python collect_findings.py <run_dir> [--from <subagent产物目录>] [--allow-missing]
"""

from __future__ import annotations
import argparse
import json
import os
import sys


def _console_safe():
    # GBK 控制台下打印 ✓/✗ 会 UnicodeEncodeError 且产物不落盘，errors=replace 自愈
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(errors="replace")
        except Exception:
            pass


VERDICTS = {"符合", "部分符合", "不符合", "不适用"}
# 5 个 subagent 分组对应的 category 名（subagent 输出应含这些 category）
GROUP_CATEGORIES = {
    "强制性条文检查", "常见设计问题检查",
    "一致性检查", "表格逻辑关系专项检查",
    "语法表述检查", "文字规范性检查",
    "设计标准检查", "法律法规检查",
}


def _read_json(path):
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _normalize_result(r):
    """校验并补齐单个 result 的字段。返回 (id, normalized_result) 或 None。"""
    if not isinstance(r, dict):
        return None
    rid = r.get("id")
    if not rid or not isinstance(rid, str):
        return None

    verdict = r.get("verdict", "未检出")
    if verdict not in VERDICTS:
        verdict = "未检出"

    # issues：list[dict]，每条只保留 4 个标准字段
    raw_issues = r.get("issues", [])
    if not isinstance(raw_issues, list):
        raw_issues = []
    issues = []
    for it in raw_issues:
        if not isinstance(it, dict):
            continue
        issues.append({
            "loc": str(it.get("loc", "")),
            "original": str(it.get("original", "")),
            "type": str(it.get("type", "")),
            "suggestion": str(it.get("suggestion", "")),
        })

    # evidence：list[dict]，每条只保留 loc/quote
    raw_ev = r.get("evidence", [])
    if isinstance(raw_ev, str):
        raw_ev = []
    elif not isinstance(raw_ev, list):
        raw_ev = []
    evidence = []
    for ev in raw_ev:
        if not isinstance(ev, dict):
            continue
        evidence.append({
            "loc": str(ev.get("loc", "")),
            "quote": str(ev.get("quote", "")),
        })

    out = {
        "id": rid,
        "name": str(r.get("name", "")),
        "verdict": verdict,
        "conclusion": str(r.get("conclusion", "")),
        "issues": issues,
        "evidence": evidence,
    }
    # 保留 _source 标注（区分 det / subagent），供 assemble_report 显示
    if "_source" in r:
        out["_source"] = r["_source"]
    return rid, out


def _load_group_files(run_dir, from_dir):
    """加载 4 个分组的 subagent 产物。返回 list[{category, results}]。"""
    search_dir = from_dir or run_dir
    groups = []
    # 优先找 groupA.json..groupD.json；也兼容历史命名 gA.json/_group_A.json
    candidates = []
    for name in ("groupA1.json", "groupA2.json",
                 "groupB.json", "groupC.json", "groupD.json",
                 "groupA.json",  # 历史运行兼容（拆组前的 A 组产物）
                 "gA.json", "gB.json", "gC.json", "gG.json",
                 "_group_A.json", "_group_B.json", "_group_C.json", "_group_D.json"):
        p = os.path.join(search_dir, name)
        if os.path.exists(p):
            candidates.append(p)
    # 也扫 run_dir 下所有 group*.json / g[ABCDEFG].json
    if os.path.isdir(search_dir):
        for fn in os.listdir(search_dir):
            if fn.endswith(".json") and (
                fn.startswith("group") or fn.startswith("g") or fn.startswith("_group")
            ) and fn not in (os.path.basename(c) for c in candidates):
                candidates.append(os.path.join(search_dir, fn))

    for p in sorted(set(candidates)):
        data = _read_json(p)
        if data is None:
            continue
        # 文件可能是单个 {category,results} 或 list[{category,results}]
        items = data if isinstance(data, list) else [data]
        for it in items:
            if isinstance(it, dict) and "results" in it:
                groups.append(it)
    return groups


def main():
    _console_safe()
    ap = argparse.ArgumentParser(description="收集并合并 subagent 产物 + det_findings")
    ap.add_argument("run_dir", help="运行产物目录")
    ap.add_argument("--from", dest="from_dir", default=None,
                    help="subagent 产物所在子目录（默认 run_dir）")
    ap.add_argument("--allow-missing", action="store_true",
                    help="确属无法补判时放行占位条目（最终报告须向用户说明）")
    args = ap.parse_args()

    run_dir = args.run_dir
    if not os.path.isdir(run_dir):
        print(f"用法: python collect_findings.py <run_dir> [--from <subagent产物目录>]")
        sys.exit(1)
    here = os.path.dirname(os.path.abspath(__file__))

    # 1) 加载 requirements（用于 name 回填 + 覆盖率校验）
    reqs = _read_json(os.path.join(here, "requirements.json")) or []
    req_by_id = {r["id"]: r for r in reqs if "id" in r}
    # category 排序（按 requirements 出现顺序）
    cat_order = []
    for r in reqs:
        if r.get("category") not in cat_order:
            cat_order.append(r["category"])

    # 2) 收集 subagent 产物
    groups = _load_group_files(run_dir, args.from_dir)
    sub_by_id = {}   # id -> result（subagent 来源）
    n_sub = 0
    for g in groups:
        for r in g.get("results", []):
            norm = _normalize_result(r)
            if norm:
                rid, res = norm
                sub_by_id[rid] = res
                n_sub += 1
    print(f"[collect] subagent 产物: {len(groups)} 个分组文件，{n_sub} 条结果")

    # 3) 加载 det_findings（优先级最高，覆盖同 id 的 subagent 结果）
    det_path = os.path.join(run_dir, "det_findings.json")
    det_data = _read_json(det_path) or []
    det_by_id = {}
    n_det = 0
    for g in det_data:
        for r in g.get("results", []):
            norm = _normalize_result(r)
            if norm:
                rid, res = norm
                res["_source"] = "deterministic"
                det_by_id[rid] = res
                n_det += 1
    print(f"[collect] det_findings: {n_det} 条直出结论（优先级最高）")

    # 4) 合并：det > subagent > 占位
    merged = {}   # id -> result
    overridden = []
    for rid, res in sub_by_id.items():
        res.setdefault("_source", "subagent")
        merged[rid] = res
    for rid, res in det_by_id.items():
        if rid in merged:
            overridden.append(rid)
        merged[rid] = res   # det 覆盖

    # name 回填 + 覆盖率校验
    missing = []
    for rid, req in req_by_id.items():
        if rid not in merged:
            missing.append(rid)
            merged[rid] = {
                "id": rid,
                "name": req.get("name", rid),
                "verdict": "未检出",
                "conclusion": "（未取得该条检查结论，建议复核。）",
                "issues": [],
                "evidence": [],
                "_source": "missing",
            }
        else:
            # 回填 name（若 subagent 漏了）
            if not merged[rid].get("name"):
                merged[rid]["name"] = req.get("name", rid)

    if overridden:
        print(f"[collect] det 覆盖 subagent: {overridden}")
    if missing:
        print(f"[collect] 缺失 {len(missing)} 条（先按“未检出”占位写入，文末门禁）: {missing}")
    else:
        print(f"[collect] ✓ 42 条全覆盖")

    # 5) 按 requirements 顺序 + category 分组输出
    out_groups = []
    cat_to_results = {}
    for r in reqs:
        rid = r["id"]
        cat_to_results.setdefault(r["category"], []).append(merged[rid])
    for cat in cat_order:
        if cat in cat_to_results:
            out_groups.append({"category": cat, "results": cat_to_results[cat]})

    out_path = os.path.join(run_dir, "findings.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out_groups, f, ensure_ascii=False, indent=1)

    # 摘要
    dist = {}
    for res in merged.values():
        v = res["verdict"]
        dist[v] = dist.get(v, 0) + 1
    print(f"[collect] 分布: {dist}")
    print(f"[collect] 已生成 {out_path}（{len(merged)} 条）")

    # 覆盖硬门禁：占位 ≠ 完成（2026-09-30 实战漏 3 条 ci_* 的教训）
    if missing:
        print()
        print(f"[collect] ✗ FAIL：仍有 {len(missing)} 条为占位（未检出/_source=missing）: {missing}")
        print("[collect] 按 SKILL.md ③ 失败恢复三梯度补齐后重跑 collect：")
        print("[collect]   1) 对失败运行 resume 收尾模式，用 json_append.py 补写缺条；")
        print("[collect]   2) 主会话 Grep 定点补判，写 groupXb.json（本脚本自动收集）；")
        print("[collect]   3) 前两步无效才重发更小任务（≤6 条/组）。")
        if args.allow_missing:
            print("[collect] --allow-missing 已指定：放行（最终报告须向用户说明未检出条目）")
        else:
            print("[collect] 确属无法补判时可用 --allow-missing 放行。")
            sys.exit(1)


if __name__ == "__main__":
    main()
