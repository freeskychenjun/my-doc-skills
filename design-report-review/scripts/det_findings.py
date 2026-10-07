# -*- coding: utf-8 -*-
"""
det_findings.py — 报告智能校审技能 · 确定性结论直出（提速 + 提质核心）

把 deterministic_checks.py 产出的、判定规则已是硬 if-else 的检查条目直接转成
findings.json 同结构，合成时并入。这些条目不再派 subagent，理由：
  1) 它们本质是数数/查表/加法，脚本比 LLM 更准；
  2) SKILL 设计原则即"宁可漏报也不误报"，det 给出的 fail 是高置信硬伤。

直出条目（共 4 条）：
  std_001  大数值单位使用规范     det_numbers.std001_bigunit.mixed_form
  std_003  流量数值保留规范       det_numbers.std003_flow.violations
  std_004  工程等别和级别数字书写 det_numbers.std004_grade
  gram_007 图表编号错误检查       det_numbering.{图,表}.{format_inconsistent,issues}

半直出条目（tbl_001/ds_001/ds_002/lr_001/lr_002）det 只能给部分证据，verdict
边界仍需 LLM，故不在本脚本直出，而是通过 slice_artifacts 把 det 摘要喂给 subagent。

用法：
  python det_findings.py <run_dir>
  → <run_dir>/det_findings.json   （findings.json 子集格式）
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

# 直出条目 id → requirements.json 里的 name（启动时从 requirements.json 校准）
DET_IDS = ["std_001", "std_003", "std_004", "gram_007"]


def _read_json(path):
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _name_map(here):
    """从 requirements.json 取 id→name 映射，保证 name 字段与要求一致。"""
    reqs = _read_json(os.path.join(here, "requirements.json")) or []
    return {r["id"]: r["name"] for r in reqs if "id" in r and "name" in r}


def _result(rid, names, verdict, conclusion, issues=None, evidence=None):
    """构造一个标准 result 对象（字段齐全，顺序固定便于人工核对）。"""
    return {
        "id": rid,
        "name": names.get(rid, rid),
        "verdict": verdict,
        "conclusion": conclusion,
        "issues": issues or [],
        "evidence": evidence or [],
        "_source": "deterministic",   # 标注机器判定，便于合成时区分
    }


# ============================================================
# std_001 大数值单位：mixed_form (亿/百万/万 与 10的N次方 混用)
# ============================================================
def build_std_001(det_numbers, names):
    bu = det_numbers.get("std001_bigunit", {})
    mixed = bu.get("mixed_form", False)
    word_forms = bu.get("word_forms", {})
    hits = bu.get("pow10_hits", [])

    if not mixed:
        wf = "、".join(f"{k}({v}处)" for k, v in word_forms.items()) if word_forms else "无"
        conclusion = (f"检查大数单位使用：文字单位 {wf}；10的次方形式 {len(hits)} 处。"
                      f"未发现文字单位与10的次方形式混用，符合规范。")
        return _result("std_001", names, "符合", conclusion)

    # mixed_form=true：存在混用——逐处列出 10的次方写法（含 ×10N 平文写法）
    issues = []
    for v in hits:
        loc = v.get("loc", "") or "全文"
        val = v.get("value", "")
        mant = v.get("mantissa", "")
        cn = v.get("cn_unit", "")
        if mant:
            sug = f"{val} → {mant}{cn}（如 0.27×104 → 0.27万）"
        else:
            sug = f"{val} → 统一改为文字单位（10的{v.get('exp')}次方 = {cn}）"
        issues.append({
            "loc": loc,
            "original": f"{val}（{v.get('context', '')[:60]}）",
            "type": "单位混用",
            "suggestion": sug,
        })
    wf = "、".join(f"{k}({v}处)" for k, v in word_forms.items())
    conclusion = (f"检查大数单位使用：文字单位 {wf}（共 {sum(word_forms.values())} 处）"
                  f"与 10的次方形式（{len(hits)} 处，含 ×10N 平文写法）混用，"
                  "违反“同一本报告中只能采用一种形式”。因全文文字单位占绝对多数，"
                  "建议把 10的次方形式逐处换算为文字单位（×10⁴＝万、×10⁸＝亿），数值不变。")
    return _result("std_001", names, "不符合", conclusion, issues)


# ============================================================
# std_003 流量数值保留：violations (sig_figs>3 或 decimals>3)
# ============================================================
def build_std_003(det_numbers, names):
    fl = det_numbers.get("std003_flow", {})
    violations = fl.get("violations", [])
    info = fl.get("info", [])
    n_v = len(violations)

    if n_v == 0:
        info_note = (f"，另有 {len(info)} 处整数流量不足3位有效数字"
                     f"（设计流量常为整数，结合语境判断，不判违规）") if info else ""
        conclusion = (f"检查流量数值有效数字：未发现>3位有效数字或小数>3位的违规"
                      f"{info_note}。")
        return _result("std_003", names, "符合", conclusion)

    # 有违规：逐条生成 issue（value→suggested）
    issues = []
    for v in violations:
        loc = v.get("para_idx", "")
        val = v.get("value", "")
        sug = v.get("suggested", "")
        sf = v.get("sig_figs", 0)
        dec = v.get("decimals", 0)
        reason = []
        if sf > 3:
            reason.append(f"{sf}位有效数字>3")
        if dec > 3:
            reason.append(f"小数{dec}位>3")
        reason_str = "、".join(reason)
        issues.append({
            "loc": loc,
            "original": f"{val}（{reason_str}）",
            "type": "有效数字",
            "suggestion": f"{val} → {sug}",
        })
    conclusion = (f"检查流量数值有效数字：发现 {n_v} 处违规"
                  f"（>3位有效数字或小数>3位），应保留3位有效数字、小数点后不超过3位。")
    return _result("std_003", names, "不符合", conclusion, issues)


# ============================================================
# std_004 工程等别/级别罗马-阿拉伯：findings 里存在 issue 项
# 判定尺度（2026-10-01）：个别笔误（≤3 处）→ 部分符合；成片误用（>3 处）→ 不符合
# ============================================================
def build_std_004(det_numbers, names):
    grades = det_numbers.get("std004_grade", [])
    issues_found = [g for g in grades if not g.get("ok") and g.get("issue")]

    if not issues_found:
        conclusion = ("检查工程等别/级别数字书写：未发现罗马数字与阿拉伯数字误用，"
                      "也未发现拉丁字母 I/V/X 冒充罗马数字的写法。")
        return _result("std_004", names, "符合", conclusion)

    issues = []
    for g in issues_found:
        issues.append({
            "loc": g.get("loc", "") or "全文（搜索匹配）",
            "original": g.get("raw", "") or g.get("loc", ""),
            "type": "数字形式",
            "suggestion": f"{g.get('rule', '')}：{g.get('issue', '')}，应改为{g.get('expect', '')}",
        })
    verdict = "部分符合" if len(issues_found) <= 3 else "不符合"
    conclusion = (f"检查工程等别/级别数字书写：发现 {len(issues_found)} 处罗马/阿拉伯"
                  f"数字书写问题（等别/航道等级/围岩及水质类别应为罗马数字Ⅰ~Ⅶ，"
                  "建筑物及堤防级别应为阿拉伯数字1~5，罗马数字须用专用字符而非拉丁字母I/V/X）。"
                  "逐处见问题清单，多数表述经扫描未见误用。")
    return _result("std_004", names, verdict, conclusion, issues)


# ============================================================
# gram_007 图表编号：format_inconsistent + issues(重复/断号)
# ============================================================
def build_gram_007(det_numbering, names):
    parts = []
    has_hard_issue = False    # 重复/断号 → 不符合
    has_format_issue = False  # 仅格式不统一 → 部分符合
    issues = []

    for t in ("图", "表"):
        blk = det_numbering.get(t, {})
        if not blk:
            continue
        count = blk.get("count", 0)
        formats = blk.get("formats", [])
        fmt_inc = blk.get("format_inconsistent", False)
        blk_issues = blk.get("issues", [])

        part = f"{t}共 {count} 个题录"
        if formats:
            part += f"，编号格式 {formats}"
        if blk_issues:
            has_hard_issue = True
            for it in blk_issues:
                issues.append({
                    "loc": f"{t}编号",
                    "original": f"{t}编号 {it.get('prefix', '')}：{it.get('detail', '')}",
                    "type": it.get("type", "编号问题"),
                    "suggestion": f"核对 {it.get('prefix', '')} 系列{t}编号，"
                                  f"修正{it.get('type', '问题')}",
                })
        elif fmt_inc:
            has_format_issue = True
            issues.append({
                "loc": f"{t}编号",
                "original": f"{t}编号格式不统一：{formats}",
                "type": "编号格式",
                "suggestion": f"统一{t}编号格式为一种（如全部用 X.Y-Z 或 X.Y.Z）",
            })
        parts.append(part)

    summary = "；".join(parts) if parts else "未扫描到图表题录"

    if has_hard_issue:
        verdict = "不符合"
        conclusion = f"图表编号检查：{summary}。发现重复编号或编号间断，需修正。"
    elif has_format_issue:
        verdict = "部分符合"
        conclusion = (f"图表编号检查：{summary}。编号连续无重复，但图/表编号格式"
                      f"存在多种写法混用，建议统一。")
    else:
        verdict = "符合"
        conclusion = f"图表编号检查：{summary}。编号连续、格式统一。"
    return _result("gram_007", names, verdict, conclusion, issues)


def main():
    args = sys.argv[1:]
    run_dir = args[0] if args else None
    if not run_dir or not os.path.isdir(run_dir):
        print("用法: python det_findings.py <run_dir>")
        sys.exit(1)
    here = os.path.dirname(os.path.abspath(__file__))
    names = _name_map(here)

    det_numbers = _read_json(os.path.join(run_dir, "det_numbers.json")) or {}
    det_numbering = _read_json(os.path.join(run_dir, "det_numbering.json")) or {}

    std_results = [
        build_std_001(det_numbers, names),
        build_std_003(det_numbers, names),
        build_std_004(det_numbers, names),
    ]
    gram_results = [build_gram_007(det_numbering, names)]

    out = [
        {"category": "文字规范性检查", "results": std_results},
        {"category": "语法表述检查", "results": gram_results},
    ]
    out_path = os.path.join(run_dir, "det_findings.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    # 摘要
    for grp in out:
        for r in grp["results"]:
            n_iss = len(r.get("issues", []))
            print(f"[det_findings] {r['id']:9s} {r['verdict']:5s} "
                  f"issues={n_iss}  ({r['name']})")
    print(f"[det_findings] 已生成 {out_path}（4 条直出结论）")


if __name__ == "__main__":
    main()
