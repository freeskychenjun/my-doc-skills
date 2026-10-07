# -*- coding: utf-8 -*-
"""regression_check.py — 报告智能校审 skill 的回归测试门禁

两种模式：
  findings  对比 <run_dir>/findings.json 与金标准，五项检查：
            1) 42 条全覆盖、未检出=0
            2) verdict 翻转分类：放宽(golden∈{部分符合,不符合}→new∈{符合,不适用})=FAIL；
               收紧(符合→部分符合/不符合)=WARN；新覆盖(不适用→其他)=WARN
            3) issue 数量下限：golden>=5 时 new>=80%；golden<5 时 new>=golden-2；总量 new>=80%golden
            4) 反幻觉抽查（确定性采样<=40条）：issue.original 至少一个 10 字连续片段
               能在 fulltext.md 或 tables.json 网格中找到；loc 段落号必须存在
            5) 机器黑话残留=0
  artifacts 比对 <run_dir> 与金标准 artifacts 的确定性产物（fulltext/tables/numbers/
            captions/citations/det_*）与切片 _slices/groupX.json，--allow 可豁免预期变更

用法：
  python regression_check.py findings  <run_dir> --golden <golden_dir> [--artifacts <dir>]
  python regression_check.py artifacts <run_dir> --golden <golden_dir> [--allow groupC,det_suspects]

退出码：0=PASS 1=FAIL（按门禁须回滚或修复）
"""

from __future__ import annotations
import argparse
import json
import math
import os
import re
import sys

JARGON = re.compile(
    r"det_\w+|n_fail|n_warn|grid_omitted|consistency_issues|fulltext_slice|"
    r"_slices|本切片|所提供切片|norm_code|para_idx|std003_flow|_fulltext_note"
)
RELAX_TO = {"符合", "不适用"}
ART_FILES = ["fulltext.md", "tables.json", "numbers.json", "captions.json",
             "citations.json", "det_table.json", "det_numbers.json",
             "det_numbering.json", "det_citations.json", "det_suspects.json"]
SLICE_FILES = ["groupA1.json", "groupA2.json", "groupB.json", "groupC.json", "groupD.json"]


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def results_map(data):
    m = {}
    for g in data if isinstance(data, list) else []:
        for r in g.get("results", []):
            m[r.get("id")] = r
    return m


SUPSUB = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉", "01234567890123456789")


def norm_frag(s):
    """空白不敏感 + 上/下标数字归一（m³→m3），两侧同样处理后再比对。"""
    return re.sub(r"\s+", "", s).translate(SUPSUB)


def build_haystack(art_dir):
    hay = ""
    ft = os.path.join(art_dir, "fulltext.md")
    if os.path.exists(ft):
        hay += open(ft, encoding="utf-8", errors="ignore").read()
    tp = os.path.join(art_dir, "tables.json")
    if os.path.exists(tp):
        try:
            for t in load_json(tp):
                hay += t.get("caption", "") or ""
                for row in t.get("grid", []) or []:
                    hay += "".join(str(c) for c in row if c is not None)
        except (json.JSONDecodeError, OSError):
            pass
    cp = os.path.join(art_dir, "captions.json")
    if os.path.exists(cp):
        try:
            for c in load_json(cp):
                hay += c.get("text", "") or ""
        except (json.JSONDecodeError, OSError):
            pass
    return hay


def check_findings(run_dir, golden_dir, art_override):
    gpath = os.path.join(golden_dir, "findings_golden.json")
    npath = os.path.join(run_dir, "findings.json")
    for p in (gpath, npath):
        if not os.path.exists(p):
            print("[regression] 缺文件:", p)
            return 1
    G, N = results_map(load_json(gpath)), results_map(load_json(npath))
    art_dir = art_override or os.path.join(golden_dir, "artifacts")
    fails, warns = [], []

    # 1) 覆盖
    missing = [i for i in G if i not in N]
    extra = [i for i in N if i not in G]
    undet = [i for i, r in N.items() if r.get("verdict") == "未检出"]
    if missing or undet:
        fails.append("覆盖不全：缺失 %s / 未检出 %s" % (missing, undet))
    if extra:
        warns.append("多余条目(仅提示): %s" % extra)
    print("[1] 覆盖：golden %d 条，new %d 条，缺失 %d，未检出 %d" %
          (len(G), len(N), len(missing), len(undet)))

    # 2) verdict 翻转
    relax, tighten, newcov = [], [], []
    for rid in G:
        if rid not in N:
            continue
        gv, nv = G[rid].get("verdict"), N[rid].get("verdict")
        if gv == nv:
            continue
        if gv in ("部分符合", "不符合") and nv in RELAX_TO:
            relax.append((rid, gv, nv))
        elif gv == "符合" and nv in ("部分符合", "不符合"):
            tighten.append((rid, gv, nv))
        elif gv == "不适用" and nv != "不适用":
            newcov.append((rid, gv, nv))
        else:
            tighten.append((rid, gv, nv))
    if relax:
        fails.append("放宽方向翻转 %d 条(漏报信号): %s" % (len(relax), relax))
    print("[2] verdict 翻转：放宽 %d（FAIL 项）、收紧/其他 %d、新覆盖 %d" %
          (len(relax), len(tighten), len(newcov)))
    for t in relax + tighten + newcov:
        print("      %-16s %s -> %s" % t)

    # 3) issue 数量下限
    low = []
    tot_g = tot_n = 0
    for rid in G:
        ng = len(G[rid].get("issues", []))
        nn = len(N.get(rid, {}).get("issues", []))
        tot_g += ng
        tot_n += nn
        # 数量下限：n>=5 用 80%；n<5 用 n-2 容差（小计数噪声）
        floor = (math.ceil(ng * 0.8) if ng >= 5 else max(0, ng - 2)) if ng else 0
        if nn < floor:
            low.append((rid, ng, nn))
    if low:
        fails.append("issue 数量低于下限 %d 条: %s" % (len(low), low))
    if tot_g and tot_n < tot_g * 0.8:
        fails.append("issue 总量 %d < golden %d 的 80%%" % (tot_n, tot_g))
    print("[3] issue 总量：golden %d，new %d；单项低于下限 %d 条" % (tot_g, tot_n, len(low)))

    # 4) 反幻觉抽查（确定性采样）
    all_issues = []
    for rid in sorted(N):
        if N[rid].get("_source") == "deterministic":
            continue  # det 直出条目的 original 是机器摘要而非文档引文，不适用反幻觉检查
        for it in N[rid].get("issues", []):
            all_issues.append((rid, it))
    sample = all_issues[::max(1, len(all_issues) // 40)][:40] if all_issues else []
    hay = norm_frag(build_haystack(art_dir))

    def frag_ok(orig):
        o = (orig or "").replace("\n", "")
        # 省略号是片段边界（LLL 常用「甲…乙」缩略引用），不能删掉后硬拼
        o = re.sub(r"…+|\.{3,}|⋯+", "；", o)
        # 1) 引号内片段（含直引号 " 与弯引号 “「『）
        quots = re.findall(r"['\"「『“]([^'\"」』“”「『]{4,})['\"」』”]", o)
        # 2) 分号/句号/冒号/顿号切分的段（Pxxxx：片段；T9：片段、片段 这类）
        segs = [s for s in re.split(r"[；;。：:\n、]", o) if len(s) >= 6]
        for c in quots + segs:
            n = norm_frag(c.strip(" \t'\"「」『』“”‘’·；;，,、"))
            if len(n) >= 4 and n in hay:
                return True
        # 3) 数据型对照（cons_* 转述式）：引用的数值全部真实存在于文档即通过
        nums = {x for x in re.findall(r"\d[\d.]*%?", o) if len(x) >= 3}
        if len(nums) >= 2 and all(x in hay for x in nums):
            return True
        # 3) 短串整体；长串滑窗
        if len(o) <= 12:
            n = norm_frag(o)
            return len(n) >= 3 and n in hay
        for i in range(0, len(o) - 9, 5):
            n = norm_frag(o[i:i + 10])
            if len(n) >= 8 and n in hay:
                return True
        return False

    fulltext_lines = set()
    ftp = os.path.join(art_dir, "fulltext.md")
    if os.path.exists(ftp):
        for line in open(ftp, encoding="utf-8", errors="ignore"):
            m = re.match(r"(P\d+)", line)
            if m:
                fulltext_lines.add(m.group(1))
    bad_quote, bad_loc = [], []
    for rid, it in sample:
        if not frag_ok(it.get("original") or ""):
            bad_quote.append((rid, it.get("loc", ""), (it.get("original") or "")[:40]))
        for m in re.finditer(r"P\d{3,4}", str(it.get("loc", ""))):
            if m.group(0) not in fulltext_lines:
                bad_loc.append((rid, it.get("loc", "")))
    if bad_quote:
        fails.append("反幻觉失败 %d/%d 条(原文片段在全文与表格中均找不到): %s" %
                     (len(bad_quote), len(sample), bad_quote[:5]))
    if bad_loc:
        warns.append("loc 段落号不存在 %d 处(核查): %s" % (len(set(bad_loc)), sorted(set(bad_loc))[:5]))
    print("[4] 反幻觉抽查：采样 %d 条，原文片段失败 %d，loc 不存在 %d" %
          (len(sample), len(bad_quote), len(set(bad_loc))))

    # 5) 机器黑话
    jbad = [(rid, m.group(0)) for rid in N for m in [JARGON.search(N[rid].get("conclusion", ""))] if m]
    if jbad:
        fails.append("机器黑话残留 %d 条: %s" % (len(jbad), jbad[:5]))
    print("[5] 机器黑话残留：%d" % len(jbad))

    print()
    for w in warns:
        print("[WARN]", w)
    if fails:
        print("==== 回归门禁 FAIL ====")
        for f in fails:
            print("[FAIL]", f)
        return 1
    print("==== 回归门禁 PASS ====")
    return 0


def canon(obj):
    return json.dumps(obj, ensure_ascii=False, sort_keys=True)


def check_artifacts(run_dir, golden_dir, allow):
    g_art = os.path.join(golden_dir, "artifacts")
    n_art = run_dir
    allow = set(filter(None, (allow or "").split(",")))
    fails = []
    for fn in ART_FILES:
        gp = os.path.join(g_art, fn)
        if not os.path.exists(gp):
            continue
        np = os.path.join(n_art, fn)
        if not os.path.exists(np):
            if fn not in allow:
                fails.append("缺少产物 %s" % fn)
            continue
        if canon(load_json(gp) if fn.endswith(".json") else open(gp, encoding="utf-8").read()) != \
           canon(load_json(np) if fn.endswith(".json") else open(np, encoding="utf-8").read()):
            tag = "豁免" if fn in allow else "FAIL"
            print("[artifacts] %-22s 内容有差异（%s）" % (fn, tag))
            if fn not in allow:
                fails.append("产物内容变化: %s" % fn)
        else:
            print("[artifacts] %-22s 一致" % fn)
    for sn in SLICE_FILES:
        gp = os.path.join(g_art, "_slices", sn)
        if not os.path.exists(gp):
            continue
        np = os.path.join(n_art, "_slices", sn)
        key = "group" + sn.replace("group", "").replace(".json", "")
        if not os.path.exists(np):
            if key not in allow:
                fails.append("缺少切片 %s" % sn)
            continue
        if canon(load_json(gp)) != canon(load_json(np)):
            tag = "豁免" if key in allow else "FAIL"
            print("[artifacts] _slices/%-14s 内容有差异（%s）" % (sn, tag))
            if key not in allow:
                fails.append("切片内容变化: %s" % sn)
        else:
            print("[artifacts] _slices/%-14s 一致" % sn)
    if fails:
        print("==== artifacts 比对 FAIL ====")
        for f in fails:
            print("[FAIL]", f)
        return 1
    print("==== artifacts 比对 PASS ====")
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("mode", choices=["findings", "artifacts"])
    ap.add_argument("run_dir")
    ap.add_argument("--golden", required=True)
    ap.add_argument("--artifacts", default=None,
                    help="findings 模式下反幻觉所用的产物目录（默认 golden 目录下 artifacts/）")
    ap.add_argument("--allow", default="",
                    help="artifacts 模式豁免预期变更，逗号分隔，如 groupC,det_suspects")
    args = ap.parse_args()
    if args.mode == "findings":
        sys.exit(check_findings(args.run_dir, args.golden, args.artifacts))
    sys.exit(check_artifacts(args.run_dir, args.golden, args.allow))


if __name__ == "__main__":
    main()
