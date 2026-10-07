# -*- coding: utf-8 -*-
"""
polish_findings.py — 报告智能校审技能 · 结论润色（去除机器内部字段名）

subagent 有时会把切片里的 det 证据字段名（如 det_citations_brief.summary、
consistency_issues=0、grid_omitted:true、n_fail、空数组 等）直接照抄进给用户
看的 conclusion，对读报告的人不友好。本脚本在 collect_findings 之后、
assemble_report 之前运行，把这些机器黑话翻译成自然语言。

清洗对象：findings.json 里每条 result 的 conclusion 字段（issue 的 suggestion
不动，因为它本就是逐字修改建议）。同时清洗 issues 的 original（仅去除字段名
残留，保留原文）。

原则：
  - 只做"黑话 → 人话"的等价替换，不改变判定含义。
  - 去除"切片""本切片"这类暴露内部实现的措辞，改为"文档/报告"。
  - 保留所有具体数值、定位、原文引用——这些是有价值的信息。

用法：
  python polish_findings.py <run_dir>
  → 原地更新 <run_dir>/findings.json
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


# ============================================================
# 黑话 → 人话 替换表（按长度降序处理，避免短串误伤长串）
# ============================================================
# 元组列表 (正则, 替换)；正则带捕获组时用函数引用
REPLACEMENTS = [
    # —— "本切片/切片提供/本组切片" → "文档/报告" ——
    (re.compile(r"本组切片"), "本文档"),
    (re.compile(r"本组文档"), "本文档"),
    (re.compile(r"本组"), "本文档"),
    (re.compile(r"所提供切片"), "所提供文档"),
    (re.compile(r"基于所提供切片"), "基于所提供文档"),
    (re.compile(r"本切片提供的"), "文档中的"),
    (re.compile(r"本切片"), "本文档"),
    (re.compile(r"切片(提供)?的"), "文档的"),
    (re.compile(r"切片(内|中)"), "文档\1"),
    (re.compile(r"切片未报"), "机器检查未发现"),
    (re.compile(r"切片(为|含|有)"), "文档\1"),

    # —— det_table_brief 相关 ——
    (re.compile(r"det_table_brief(对T|在T|对|在|可见的)"), r"机器核算结果\1"),
    (re.compile(r"det_table_brief"), "机器核算结果"),
    (re.compile(r"未提供det核算结果"), "未做机器核算"),
    (re.compile(r"n_fail\s*=\s*0\s*,\s*n_warn\s*=\s*0"), "0 项不合格、0 项存疑"),
    (re.compile(r"n_fail\s*=\s*(\d+)"), r"\1 项不合格"),
    (re.compile(r"n_warn\s*=\s*(\d+)"), r"\1 项存疑"),
    (re.compile(r"n_checks"), "检查项"),
    (re.compile(r"(\d+)项检查均未报fail或warn"), r"\1 项检查均无不合理"),
    (re.compile(r"未报fail或warn"), "未发现不合理"),
    (re.compile(r"无fail/warn"), "未发现不合理"),
    (re.compile(r"报fail或warn"), "发现不合理"),
    (re.compile(r"info级"), "提示级"),
    (re.compile(r"info 项"), "提示项"),
    (re.compile(r"info项"), "提示项"),
    (re.compile(r"\binfo\b"), "提示"),
    # —— stated/computed 组合（stated最大值/computed最大值 这类）——
    (re.compile(r"stated最大值"), "所列最大值"),
    (re.compile(r"computed最大值"), "实算最大值"),
    (re.compile(r"stated\s*([\d.])"), r"所列值\1"),
    (re.compile(r"computed\s*([\d.])"), r"实算值\1"),

    # —— grid_omitted / grid 相关 ——
    (re.compile(r"grid_omitted\s*:\s*true"), "未提供单元格数据"),
    (re.compile(r"grid_omitted"), "未提供单元格数据"),
    (re.compile(r"缺乏grid(的情况)?(下)?，?"), "未提供表格单元格数据，"),
    (re.compile(r"提供完整grid"), "提供完整表格数据"),

    # —— det_citations_brief 相关 ——
    (re.compile(r"det_citations_brief\.summary\s*显示\s*"), ""),
    (re.compile(r"det_citations_brief\.summary\s*"), ""),
    (re.compile(r"det_citations_brief\.consistency\s*为空数组"), "未发现前后写法不一致"),
    (re.compile(r"det_citations_brief\.consistency"), "一致性检查"),
    (re.compile(r"det_citations_brief"), "机器核对结果"),

    # —— consistency / consistency_issues 相关 ——
    (re.compile(r"consistency_issues\s*=\s*0"), "一致性检查未发现问题"),
    (re.compile(r"consistency_issues\s*=\s*(\d+)"), r"一致性检查发现 \1 处不一致"),
    (re.compile(r"consistency\s*为空数组"), "未发现前后写法不一致"),
    (re.compile(r"\bconsistency\b"), "一致性检查"),

    # —— summary / 计数字段 ——
    (re.compile(r"law_ref_total\s*=\s*(\d+)"), r"共引用法律法规 \1 项"),
    (re.compile(r"law_not_in_list\s*=\s*(\d+)"), r"其中 \1 项不在现行法律法规清单"),
    (re.compile(r"std_total\s*=\s*(\d+)"), r"共引用标准 \1 项"),
    (re.compile(r"std_valid\s*=\s*(\d+)"), r"其中 \1 项现行有效"),
    (re.compile(r"std_not_in_list\s*=\s*(\d+)"), r"\1 项不在现行标准清单"),
    (re.compile(r"std_year_mismatch\s*=\s*(\d+)"), r"\1 项年份过期"),
    (re.compile(r"name_code_mismatch\s*=\s*(\d+)"), r"\1 项名称与编号不符"),
    (re.compile(r"name_not_in_list\s*=\s*(\d+)"), r"\1 项标准名称不在清单"),

    # —— det_std003_brief 等其他 det 字段 ——
    (re.compile(r"det_std003_brief\s*显示\s*"), ""),
    (re.compile(r"det_std003_brief"), "机器核对"),
    (re.compile(r"det_numbering"), "编号核对"),
    (re.compile(r"det_citations"), "引用核对"),
    (re.compile(r"det_numbers"), "数值检查"),
    (re.compile(r"det_table"), "表格机器核算"),
    (re.compile(r"det_findings"), "机器直出结论"),
    (re.compile(r"det_suspects"), "机器疑点候选"),
    (re.compile(r"std003_flow"), "流量有效数字检查"),
    (re.compile(r"std001_bigunit"), "大数单位检查"),
    (re.compile(r"std004_grade"), "等别级别检查"),

    # —— references_index / _slices / groupX ——
    (re.compile(r"references_index(_hits)?"), "标准/法规清单"),
    (re.compile(r"_slices/[a-zA-Z0-9_]+\.json"), "校审数据"),
    (re.compile(r"group[A-D]\.json"), "校审数据"),

    # —— not_in_list / validity 等状态值 ——
    (re.compile(r"not_in_list"), "不在现行清单"),
    (re.compile(r"year_mismatch"), "年份过期"),
    (re.compile(r"name_code_mismatch"), "名称与编号不符"),
    (re.compile(r"validity"), "有效性"),

    # —— 其他技术词 ——
    (re.compile(r"para_idx"), "段落"),
    (re.compile(r"caption_num"), "表号"),
    (re.compile(r"\bstated\b\s*"), "所列"),
    (re.compile(r"\bcomputed\b\s*"), "实算"),
    (re.compile(r"\bpass\b"), "通过"),
    (re.compile(r"\bfail\b"), "不合格"),
    (re.compile(r"\bwarn\b"), "存疑"),

    # —— 清理多余空格、重复标点（替换后产生的） ——
    (re.compile(r"，[，、\s]*，"), "，"),
    (re.compile(r"。\s*，"), "。"),
    (re.compile(r"[，；]\s*。"), "。"),
    (re.compile(r"\(\s*含名称单独引用\s*std_name_only\s*条目\s*\)"), "（含仅以名称引用的条目）"),
    (re.compile(r"std_name_only"), "仅以名称引用"),
    # —— 替换后产生的语义重复/不通顺 ——
    (re.compile(r"机器核算结果（机器核算结果）"), "机器核算结果"),
    (re.compile(r"机器核算结果结果"), "机器核算结果"),
    (re.compile(r"机器核算结果(\s*)对T"), r"机器核算\1对表 T"),
    (re.compile(r"机器核算结果(\s*)在T"), r"机器核算\1对表 T"),
    (re.compile(r"机器核算结果(\s*)T0"), r"机器核算\1对表 T0"),
    (re.compile(r"机器核算(\s*)T0"), r"机器核算对表 T0"),
    (re.compile(r"均标记未提供单元格数据（未含单元格数据）"), "均未提供单元格数据"),
    (re.compile(r"项\s+不在现行"), "项不在现行"),
    (re.compile(r"，且\s*未发现前后写法不一致"), "，前后写法一致"),
    (re.compile(r"未发现前后写法不一致。同一标准"), "前后写法一致。同一标准"),
    (re.compile(r"未发现前后写法不一致。文档引用"), "前后写法一致。文档引用"),
    # —— LLM 原文措辞残留（清洗后产生的不通顺）——
    (re.compile(r"含名称单独引用\s+仅以名称引用\s*条目"), "含仅以名称引用的条目"),
    (re.compile(r"在未提供表格单元格数据，"), "由于未提供表格单元格数据，"),
    (re.compile(r"提示级提示"), "提示级数据"),
]


def polish_text(text: str) -> str:
    """对一段文本应用全部替换规则。"""
    if not text:
        return text
    out = text
    for pat, rep in REPLACEMENTS:
        out = pat.sub(rep, out)
    # 收尾：清理连续空格、句首多余标点、空括号
    out = re.sub(r"[ \t]{2,}", " ", out)
    out = re.sub(r"^[，、；\s]+", "", out, flags=re.MULTILINE)
    out = re.sub(r"\(\s*\)", "", out)          # 空括号
    out = re.sub(r"（\s*）", "", out)           # 空全角括号
    # 再扫一遍主表，处理"替换后新产生的字段名"（如第一轮把某串拆出 det_ 残片）
    for pat, rep in REPLACEMENTS:
        out = pat.sub(rep, out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def polish_result(r: dict) -> bool:
    """润色单条 result 的 conclusion 和 issues 的 original。返回是否有改动。"""
    changed = False
    if "conclusion" in r:
        new = polish_text(r["conclusion"])
        if new != r["conclusion"]:
            r["conclusion"] = new
            changed = True
    for it in r.get("issues", []):
        if isinstance(it, dict) and "original" in it:
            new = polish_text(it["original"])
            if new != it["original"]:
                it["original"] = new
                changed = True
    return changed


def main():
    args = sys.argv[1:]
    run_dir = args[0] if args else None
    if not run_dir or not os.path.isdir(run_dir):
        print("用法: python polish_findings.py <run_dir>")
        sys.exit(1)
    fp = os.path.join(run_dir, "findings.json")
    if not os.path.exists(fp):
        print(f"[polish] 未找到 {fp}，请先运行 collect_findings.py")
        sys.exit(1)

    data = json.load(open(fp, encoding="utf-8"))
    n_changed = 0
    changed_ids = []
    for grp in data:
        for r in grp.get("results", []):
            if polish_result(r):
                n_changed += 1
                changed_ids.append(r.get("id", "?"))

    # 写回（保留 det_findings 的 _source 标注，不动）
    json.dump(data, open(fp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    if n_changed:
        print(f"[polish] 已润色 {n_changed} 条结论: {changed_ids}")
    else:
        print("[polish] 未发现需要润色的机器字段名")
    print(f"[polish] 已更新 {fp}")


if __name__ == "__main__":
    main()
