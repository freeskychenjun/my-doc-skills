# -*- coding: utf-8 -*-
"""
deterministic_checks.py — 基于抽取产物产出"机器可验证"证据。

读取 <run_dir>/ 下的 tables.json, captions.json, citations.json, numbers.json, fulltext.md
以及 scripts/references_index.json，输出：
  det_table.json      表格统计值核算（行/列合计、小计、平均、百分比、极值）
  det_citations.json  标准/法规引用：格式、有效性、名称-编号一致性、前后一致性
  det_numbers.json    std_001 大数单位混用 / std_003 流量有效数字 / std_004 等别级别数字
  det_numbering.json  gram_007 图表编号顺序/格式

证据供语义检查 subagent 引用，保证数值类、对照类检查的严谨与高效。
设计原则：宁可漏报(留给语义判断)也不误报，避免污染校审结论的可信度。
"""

from __future__ import annotations
import json
import math
import os
import re
import sys

# GBK 控制台自愈：打印 ✓/✗ 等会 UnicodeEncodeError 且产物不落盘（2026-09-30 实测）
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib_docx import norm_code, norm_name  # noqa: E402

INDEX = None  # references_index.json


# ---------------- 通用数值解析 ----------------
def parse_cell(s):
    """解析单元格：返回 (kind, value)。kind: 'num'|'pct'|'range'|'none'。"""
    if s is None:
        return ("none", None)
    t = str(s).strip()
    if t == "" or t in ("—", "-", "－", "/", "无", "略"):
        return ("none", None)
    is_pct = "%" in t
    clean = re.sub(r"[,%‰]", "", t)
    clean = re.sub(r"[a-zA-Z㎡m³²·/\s]", "", clean)
    clean = clean.replace("，", "").replace("。", "")
    m = re.match(r"^([0-9]+(?:\.[0-9]+)?)\s*[～\-~至到]\s*([0-9]+(?:\.[0-9]+)?)$", clean)
    if m:
        return ("range", (float(m.group(1)), float(m.group(2))))
    m = re.match(r"^([0-9]+(?:\.[0-9]+)?)$", clean)
    if m:
        return ("pct" if is_pct else "num", float(m.group(1)))
    m = re.match(r"^([0-9]+(?:\.[0-9]+)?)", clean)
    if m:
        return ("pct" if is_pct else "num", float(m.group(1)))
    return ("none", None)


def sig_figs(x: float, raw: str = None):
    """计算有效数字位数。

    水文流量按 SL/T 247《水文资料整编规范》取 3 位有效数字。关键规则：
      - 整数末尾的零（如 2050 末位的 0）在水文整编语境下是修约占位符，
        不算有效数字——2050 是"修约到 3 位有效数字"的结果（2,0,5 有效，
        末位 0 占位），故算 3 位。
      - 小数末尾的零（如 12.0 末位的 0）是有效的，算入有效数字——
        12.0 是 3 位（1,2,0），体现修约到 3 位有效数字。
      - 科学计数法（如 2.05e3）系数部分全算有效数字。
      - 前导零（0.50 的 0.）不算有效数字——0.50 是 2 位（5,0）。
    raw: 原始字符串，用于区分整数末尾零 vs 小数末尾零。
    """
    if x == 0:
        return 1
    src = raw.strip() if raw else ("%g" % x)
    # 去符号
    src = src.lstrip("+-")
    # 科学计数法：取系数部分
    if "e" in src.lower():
        src = src.lower().split("e")[0]
    # 拆整数/小数部分
    if "." in src:
        int_part, dec_part = src.split(".", 1)
    else:
        int_part, dec_part = src, ""
    # 整数部分：去前导零
    int_part = int_part.lstrip("0")
    if not int_part and not dec_part:
        return 1
    if "." in src:
        # 含小数点：整数部分（去前导零后）+ 小数部分全部数字（含末尾零）都有效
        n = len(int_part) + len(dec_part)
        # 但整数部分为 0 时（如 0.50），整数部分那个 0 不算
        if not int_part:
            n = len(dec_part)
        return max(n, 1)
    else:
        # 纯整数：去前导零后的位数 - 末尾连续零（水文修约占位，不算有效）
        stripped = int_part.rstrip("0")
        return max(len(stripped), 1)


_CLEAN_NUM_RE = re.compile(r"^[+-]?[0-9]+(?:\.[0-9]+)?%?$")


def is_clean_num(cell):
    """单元格是否为"纯数值"（可安全参与求和）。
    排除 '1007/27'(kW/台双单位)、'0.50～0.65'(区间)、'17.9m'(带单位) 等。"""
    if cell is None:
        return False
    t = str(cell).strip().replace(" ", "")
    return bool(_CLEAN_NUM_RE.match(t))


# ---------------- 表格统计核算 ----------------
# 严格判定，避免把"最大24h降雨""经验频率"等术语误判为统计行/占比列。
PCT_COMPOSITION_HDR = re.compile(r"(占比|比例|百分比|百分率|构成|所占|权重比)")
PCT_EXCLUDE_HDR = re.compile(r"(频率|经验频率|重现期|保证率|累积|累计|概率|超过|超越|次频率)")


def _row_label(grid, ri, ncols):
    parts = []
    for ci in range(min(2, ncols)):
        v = grid[ri][ci] if ci < len(grid[ri]) else None
        if v:
            parts.append(str(v))
    return "".join(parts).strip()


def _row_role(label_text):
    """行角色：total/subtotal/avg/max/min/data。要求标签基本等于关键词，避免误匹配。"""
    t = label_text.replace(" ", "")
    if len(t) > 12:
        return "data"
    if re.search(r"(合\s*计|总\s*计)", t) and not re.search(r"小\s*计", t):
        return "total"
    if re.search(r"小\s*计", t):
        return "subtotal"
    if re.search(r"(平\s*均|均\s*值|加权|多年平均|年平均)", t):
        return "avg"
    if re.match(r"^(最\s*大\s*值|最\s*大|最\s*高)", t):
        return "max"
    if re.match(r"^(最\s*小\s*值|最\s*小|最\s*低)", t):
        return "min"
    return "data"


def check_tables(tables):
    out = []
    for tb in tables:
        grid = tb["grid"]
        nrows = tb["nrows"]; ncols = tb["ncols"]
        if nrows < 2 or ncols < 2:
            continue
        mat, clean, roles = [], [], []
        for ri, r in enumerate(grid):
            row_v, row_c = [], []
            for c in r:
                k, v = parse_cell(c)
                row_v.append(v if k in ("num", "pct") else None)
                row_c.append(v if (is_clean_num(c) and k in ("num",)) else None)
            mat.append(row_v); clean.append(row_c)
            roles.append("header" if ri == 0 else _row_role(_row_label(grid, ri, ncols)))
        header = grid[0] if grid else []

        def near(a, b, rel=0.01, abs_=0.05):
            if a is None or b is None:
                return None
            return abs(a - b) <= max(abs_, rel * max(abs(a), abs(b), 1e-9))

        def hdr_str(ci):
            return str(header[ci]) if ci < len(header) and header[ci] else ""

        checks = []
        data_rows = [ri for ri in range(1, nrows) if roles[ri] == "data"]
        total_rows = [ri for ri in range(1, nrows) if roles[ri] == "total"]
        multi_total_rows = len(total_rows) > 1
        multi_total_cols = len([ci for ci in range(ncols)
                                if re.search(r"(合\s*计|总\s*计)", hdr_str(ci))
                                and not re.search(r"小\s*计", hdr_str(ci))]) > 1

        def is_index_col(ci):
            if re.search(r"(序号|编号|项次|项别|序次)", hdr_str(ci)):
                return True
            vals = [clean[rj][ci] for rj in data_rows if clean[rj][ci] is not None]
            if len(vals) >= 4 and vals == list(range(1, len(vals) + 1)):
                return True
            return False

        # ---- 行级：合计/总计/最大/最小 ----
        for ri in range(1, nrows):
            role = roles[ri]
            data_cols = [ci for ci in range(ncols) if mat[ri][ci] is not None]
            if role == "total":
                for ci in data_cols:
                    stated = clean[ri][ci]
                    if stated is None:
                        continue
                    comp_cells = [clean[rj][ci] for rj in data_rows]
                    if sum(1 for x in comp_cells if x is None) > 1:
                        continue  # 非纯数值过多(如kW/台双单位)，跳过避免误判
                    comp = sum(x for x in comp_cells if x is not None)
                    ok = near(stated, comp, rel=0.01, abs_=0.5)
                    if multi_total_rows:
                        checks.append({"type": "行合计(信息-多合计行未定分组)",
                                       "loc": f"R{ri}C{ci}", "stated": stated,
                                       "computed": round(comp, 4), "status": "info",
                                       "note": f"列='{hdr_str(ci)[:14]}' 供语义核查"})
                    elif ok is False:
                        checks.append({"type": "行合计", "loc": f"R{ri}C{ci}",
                                       "stated": stated, "computed": round(comp, 4),
                                       "status": "fail",
                                       "note": f"列='{hdr_str(ci)[:14]}' 差额{stated - comp:+.2f}"})
                    elif ok:
                        checks.append({"type": "行合计", "loc": f"R{ri}C{ci}",
                                       "stated": stated, "computed": round(comp, 4),
                                       "status": "pass", "note": f"列='{hdr_str(ci)[:14]}'"})
            elif role in ("max", "min"):
                for ci in data_cols:
                    stated = clean[ri][ci]
                    if stated is None:
                        continue
                    col_vals = [clean[rj][ci] for rj in data_rows if clean[rj][ci] is not None]
                    if len(col_vals) < 2:
                        continue
                    comp = max(col_vals) if role == "max" else min(col_vals)
                    checks.append({"type": ("最大值" if role == "max" else "最小值") + "(信息)",
                                   "loc": f"R{ri}C{ci}", "stated": stated, "computed": comp,
                                   "status": "info",
                                   "note": f"列='{hdr_str(ci)[:14]}' 供语义核查"})
            elif role == "avg":
                for ci in data_cols:
                    stated = mat[ri][ci]
                    col_vals = [mat[rj][ci] for rj in data_rows if mat[rj][ci] is not None]
                    if not col_vals:
                        continue
                    comp = sum(col_vals) / len(col_vals)
                    checks.append({"type": "行平均(信息)", "loc": f"R{ri}C{ci}",
                                   "stated": stated, "computed": round(comp, 4),
                                   "status": "info",
                                   "note": f"列='{hdr_str(ci)[:14]}' 偏差{(stated - comp):+.2f}"})

        # ---- 列级：表头含 合计/总计 ----
        total_cols = [ci for ci in range(ncols)
                      if re.search(r"(合\s*计|总\s*计)", hdr_str(ci)) and not re.search(r"小\s*计", hdr_str(ci))]
        for ci in total_cols:
            other_data = [cj for cj in range(ncols)
                          if cj != ci and cj not in total_cols and not is_index_col(cj)]
            for ri in data_rows:
                stated = clean[ri][ci]
                if stated is None:
                    continue
                vals = [clean[ri][cj] for cj in other_data if clean[ri][cj] is not None]
                if not vals:
                    continue
                comp = sum(vals)
                ok = near(stated, comp, rel=0.01, abs_=0.5)
                if multi_total_cols:
                    checks.append({"type": "列合计(信息-多合计列)", "loc": f"R{ri}C{ci}",
                                   "stated": stated, "computed": round(comp, 4),
                                   "status": "info", "note": "供语义核查"})
                elif ok is False:
                    checks.append({"type": "列合计", "loc": f"R{ri}C{ci}",
                                   "stated": stated, "computed": round(comp, 4),
                                   "status": "fail", "note": f"差额{stated - comp:+.2f}"})
                elif ok:
                    checks.append({"type": "列合计", "loc": f"R{ri}C{ci}",
                                   "stated": stated, "computed": round(comp, 4),
                                   "status": "pass"})

        # ---- 占比/百分比列求和（仅构成类，排除频率类）----
        for ci in range(ncols):
            hdr = hdr_str(ci)
            if PCT_EXCLUDE_HDR.search(hdr) or not PCT_COMPOSITION_HDR.search(hdr):
                continue
            cnt = sum(1 for ri in data_rows if mat[ri][ci] is not None)
            if cnt < 3:
                continue
            vals = [mat[ri][ci] for ri in data_rows if mat[ri][ci] is not None]
            s = sum(vals)
            status = "pass" if abs(s - 100) <= 2 else ("warn" if abs(s - 100) <= 5 else "fail")
            checks.append({"type": "百分比求和", "loc": f"C{ci} '{hdr[:14]}'",
                           "stated": 100, "computed": round(s, 4),
                           "status": status, "note": f"列共{cnt}个数值，合计{s:.2f}"})

        if checks:
            fails = [c for c in checks if c["status"] == "fail"]
            warns = [c for c in checks if c["status"] == "warn"]
            out.append({
                "id": tb["id"], "pid": tb["pid"], "caption": tb["caption"],
                "section": tb["section"], "nrows": nrows, "ncols": ncols,
                "header": tb["header"][:8],
                "n_checks": len(checks), "n_fail": len(fails), "n_warn": len(warns),
                "checks": checks,
            })
    return out


# ---------------- 引用检查 ----------------
def check_citations(citations, index):
    std_by_code = index["standards_by_code"]
    std_by_name = index["standards_by_name"]
    law_by_name = index["laws_by_name"]
    prefix_idx = index["names_by_prefix_code"]

    def prefix_num(nc):
        return re.sub(r"-\d{4}$", "", nc)

    results = []
    for c in citations:
        kind = c["kind"]
        rec = {"kind": kind, "raw": c["raw"], "name": c.get("name", ""),
               "para_idx": c["para_idx"], "section": c["section"],
               "context": c.get("context", "")}
        if kind == "std":
            nc = c["norm_code"]
            rec["norm_code"] = nc
            if nc in std_by_code:
                rec["validity"] = "valid"
                rec["matched_code"] = nc
            else:
                cand = prefix_idx.get(prefix_num(nc))
                if cand:
                    rec["validity"] = "year_mismatch"
                    rec["matched_code"] = cand[0]["编号"]
                    rec["note"] = f"清单中该编号现行版本：{cand[0]['编号']}《{cand[0]['名称']}》(实施{cand[0]['实施日期']})"
                else:
                    rec["validity"] = "not_in_list"
                    rec["note"] = "未在水利标准有效清单中，疑似已废止/编号有误/非水利行业"
            nm = c.get("norm_name", "")
            if nm:
                if nm in std_by_name:
                    name_idxs = set(std_by_name[nm])
                    code_idxs = set(std_by_code.get(nc, []))
                    # 仅当"名称对应序号"与"编号对应序号"完全不相交时才算名称-编号不一致
                    if code_idxs and name_idxs.isdisjoint(code_idxs):
                        rec["name_code_mismatch"] = True
                        rec["note2"] = f"名称《{c.get('name')}》清单对应 {sorted(name_idxs)}，与编号 {nc}({sorted(code_idxs)}) 不一致"
                else:
                    rec["name_not_in_list"] = True
                    rec["note2"] = f"名称《{c.get('name')}》未在标准清单中，可能名称有误"
            fmt = []
            if "—" in c["raw"] or "－" in c["raw"]:
                fmt.append("破折号为全角(建议统一)")
            if re.match(r"^(SL|GB|NB|DL|JTG|JTS|CJJ|DB)\d",
                        c["raw"].upper().replace(" ", "")):
                fmt.append("前缀与编号间无空格(格式)")
            rec["format_issues"] = fmt
            results.append(rec)
        elif kind in ("law", "law_code"):
            nm = c.get("norm_name", "")
            rec["norm_name"] = nm
            if nm and nm in law_by_name:
                rec["validity"] = "valid"
                rec["matched"] = law_by_name[nm]
            elif nm:
                rec["validity"] = "not_in_list"
                rec["note"] = f"《{c.get('name')}》未在法律法规清单中"
            else:
                rec["validity"] = "code_only"
            results.append(rec)
        elif kind == "std_name_only":
            nm = c.get("norm_name", "")
            rec["norm_name"] = nm
            if nm in std_by_name:
                rec["validity"] = "valid_name"
                rec["matched"] = std_by_name[nm]
            else:
                rec["validity"] = "not_in_list"
                rec["note"] = f"《{c.get('name')}》未在水利标准清单中"
            results.append(rec)

    # ---- 一致性 ----
    consistency = []
    grp = defaultdict(set); grp_ctx = defaultdict(list)
    for r in results:
        if r.get("norm_code") and r.get("validity") in ("valid", "year_mismatch"):
            pn = prefix_num(r["norm_code"])
            grp[pn].add(r["raw"]); grp_ctx[pn].append(r)
    for pn, raws in grp.items():
        norms = set(norm_code(x) for x in raws)
        if len(norms) > 1:
            consistency.append({"type": "同一标准编号前后写法不一致", "key": pn,
                                "variants": sorted(norms),
                                "samples": [f"{r['raw']}({r['para_idx']})" for r in grp_ctx[pn][:6]]})
    ngrp = defaultdict(set); nctx = defaultdict(list)
    for r in results:
        nm = r.get("norm_name") or r.get("name")
        if nm and r.get("norm_code"):
            ngrp[nm].add(prefix_num(r["norm_code"])); nctx[nm].append(r)
    for nm, codes in ngrp.items():
        if len(codes) > 1:
            consistency.append({"type": "同一名称对应不同编号", "key": nm,
                                "codes": sorted(codes),
                                "samples": [f"{r['raw']}({r['para_idx']})" for r in nctx[nm][:6]]})

    std_results = [r for r in results if r["kind"] == "std"]
    law_results = [r for r in results if r["kind"] in ("law", "law_code", "std_name_only")
                   and r["raw"].startswith("《")]
    return {
        "citations": results, "consistency": consistency,
        "summary": {
            "std_total": len(std_results),
            "std_valid": sum(1 for r in std_results if r.get("validity") == "valid"),
            "std_not_in_list": sum(1 for r in std_results if r.get("validity") == "not_in_list"),
            "std_year_mismatch": sum(1 for r in std_results if r.get("validity") == "year_mismatch"),
            "name_code_mismatch": sum(1 for r in results if r.get("name_code_mismatch")),
            "name_not_in_list": sum(1 for r in results if r.get("name_not_in_list")),
            "law_ref_total": len(law_results),
            "law_not_in_list": sum(1 for r in law_results if r.get("validity") == "not_in_list"),
            "consistency_issues": len(consistency),
        }}


# ---------------- 数值规范性 ----------------
def _to_3sig(x):
    if x == 0:
        return "0"
    d = 3 - int(math.floor(math.log10(abs(x)))) - 1
    d = max(d, 0)
    return ("%." + str(d) + "f") % round(x, d)


def check_numbers(numbers, fulltext):
    # std_003 流量有效数字（仅判 >3位有效数字 或 小数>3位 为硬违规）
    flows = [n for n in numbers if n["key"] == "flows"]
    flow_violations, flow_info = [], []
    seen = set()
    for n in flows:
        m = re.search(r"([0-9]+(?:\.[0-9]+)?)", n["raw"])
        if not m:
            continue
        numstr = m.group(1)
        if numstr in seen:
            continue
        seen.add(numstr)
        try:
            val = float(numstr)
        except ValueError:
            continue
        decimals = len(numstr.split(".")[1]) if "." in numstr else 0
        sf = sig_figs(val, numstr)
        if sf > 3 or decimals > 3:
            flow_violations.append({"value": numstr, "sig_figs": sf, "decimals": decimals,
                                    "suggested": _to_3sig(val), "context": n["context"],
                                    "para_idx": n["para_idx"]})
        elif sf < 3 and "." not in numstr:
            flow_info.append({"value": numstr, "sig_figs": sf, "context": n["context"],
                              "para_idx": n["para_idx"], "note": "整数不足3位有效数字(设计流量常为整数,结合语境判断)"})

    # std_001 大数单位：仅在"文字单位"与"10的N次方"两种形式混用时才算违规
    bigunits = [n for n in numbers if n["key"] == "bigunits"]
    bu_count = defaultdict(int)
    for n in bigunits:
        m = re.search(r"(亿|万|百万)", n["raw"])
        if m:
            bu_count[m.group(1)] += 1
    word_forms = {k: v for k, v in bu_count.items() if v > 0}
    has_word = len(word_forms) > 0
    # 10的次方形式直接扫全文（2026-10-01 修：extract 的 pow10 分类只认"10的N次方"
    # 字面文本，漏掉"×104"平文写法——山许实测漏 25 处导致 std_001 误判符合）
    pow10_hits = _scan_pow10(fulltext)
    has_pow10 = len(pow10_hits) > 0
    std001_mixed = has_word and has_pow10  # 同一报告只能用一种"形式"（文字 vs 10的次方）

    return {
        "std003_flow": {"violations": flow_violations, "info": flow_info,
                        "n_violations": len(flow_violations)},
        "std001_bigunit": {"word_forms": word_forms, "pow10_count": len(pow10_hits),
                           "pow10_hits": pow10_hits,
                           "mixed_form": std001_mixed,
                           "note": "亿/百万/万属同一(文字)形式可并用；仅当与'10的N次方'混用时违规"},
        "std004_grade": _check_grade(fulltext),
    }


# ×10N 平文写法（如 0.27×104t）：Word 上标丢失后即为该形式，与文字单位混用时违规。
# 指数限定 3~12（"50×100"之类量纲乘积、页码编号不误收）；同时收"10的N次方"字面写法。
POW10_CN = {3: "千", 4: "万", 5: "十万", 6: "百万", 7: "千万", 8: "亿",
            9: "十亿", 10: "百亿", 11: "千亿", 12: "万亿"}


def _scan_pow10(fulltext):
    hits = []
    pat_x = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*[×x*]\s*10(\d{1,2})(?![\d.])")
    pat_txt = re.compile(r"10\s*的\s*(\d{1,2})\s*次方")
    for line in fulltext.splitlines():
        mloc = re.match(r"\s*(P\d+)", line)
        loc = mloc.group(1) if mloc else ""
        for m in pat_x.finditer(line):
            exp = int(m.group(2))
            if exp not in POW10_CN:
                continue
            s, e = m.span()
            hits.append({"value": re.sub(r"\s+", "", m.group(0)),
                         "mantissa": m.group(1), "exp": exp, "cn_unit": POW10_CN[exp],
                         "context": line[max(0, s - 25):min(len(line), e + 25)].strip(),
                         "loc": loc})
        for m in pat_txt.finditer(line):
            exp = int(m.group(1))
            if exp not in POW10_CN:
                continue
            s, e = m.span()
            hits.append({"value": re.sub(r"\s+", "", m.group(0)),
                         "mantissa": "", "exp": exp, "cn_unit": POW10_CN[exp],
                         "context": line[max(0, s - 20):min(len(line), e + 20)].strip(),
                         "loc": loc})
    return hits


def _check_grade(fulltext):
    """std_004 等别/级别数字书写。按行扫描便于给出 P 段落定位；规则力求高精度：
    等别/航道等级/围岩类别/水质类别=罗马数字专用字符，建筑物级别=阿拉伯数字。
    2026-10-01 扩充：原版只认"工程等别/建筑物级别"紧邻窗口，漏掉"III等"（拉丁字母
    冒充罗马数字，山许 P0285 实例）等情形。"""
    findings = []

    # 拉丁字母 I/V/X 序列冒充罗马数字（"III等"应为"Ⅲ等"；单字母不收——"V型""X型"合法）
    pat_fake_roman = re.compile(r"(?<![A-Za-z])([IVX]{2,4})(?=[等级类型])")
    # 建筑物/堤防级别误用罗马数字（"Ⅲ级堤防"应为"3级堤防"；阶地/航道用罗马属正确，不收）
    pat_roman_grade = re.compile(r"([ⅠⅡⅢⅣⅤⅥⅦ])\s*级(?=[建筑物堤防])")
    # 围岩/地表水/水质类别误用阿拉伯数字（"地表水3类"应为"Ⅲ类"）
    pat_arabic_class = re.compile(r"(围岩|地表水|水质|水环境)[^，。；\n]{0,6}?([1-5])\s*类")
    # 航道/船闸等级误用阿拉伯数字（"3级航道"应为"Ⅲ级航道"）
    pat_arabic_channel = re.compile(r"(?:[1-7])\s*级(?=航道|船闸)")

    for line in fulltext.splitlines():
        mloc = re.match(r"\s*(P\d+)", line)
        loc = mloc.group(1) if mloc else ""
        for m in re.finditer(r"(枢纽工程|工程)等别[^，。；\n]{0,18}", line):
            seg = m.group(0)
            has_roman = bool(re.search(r"[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]", seg))
            arabic = re.findall(r"[1-7]", seg)
            if arabic and not has_roman:
                findings.append({"rule": "工程等别应为罗马数字(Ⅰ~Ⅴ)", "loc": loc,
                                 "raw": seg, "issue": f"出现阿拉伯数字 {arabic}",
                                 "expect": "罗马数字"})
        for m in re.finditer(r"(主要建筑物|次要建筑物|水工建筑物|建筑物|堤防)级别[^，。；\n]{0,12}", line):
            seg = m.group(0)
            roman = re.findall(r"[ⅠⅡⅢⅣⅤⅥⅦ]", seg)
            if roman:
                findings.append({"rule": "建筑物级别应为阿拉伯数字(1~5)", "loc": loc,
                                 "raw": seg, "issue": f"出现罗马数字 {roman}",
                                 "expect": "阿拉伯数字"})
        for m in pat_fake_roman.finditer(line):
            s = m.start()
            findings.append({"rule": "罗马数字应为专用字符(Ⅰ~Ⅻ)，非拉丁字母I/V/X",
                             "loc": loc, "raw": line[max(0, s - 15):s + 20].strip(),
                             "issue": f"拉丁字母 {m.group(1)} 冒充罗马数字",
                             "expect": "对应专用罗马数字字符"})
        for m in pat_roman_grade.finditer(line):
            s = m.start()
            findings.append({"rule": "建筑物/堤防级别应为阿拉伯数字(1~5)",
                             "loc": loc, "raw": line[max(0, s - 15):s + 20].strip(),
                             "issue": f"罗马数字 {m.group(1)} 用于级别",
                             "expect": "阿拉伯数字"})
        for m in pat_arabic_class.finditer(line):
            s = m.start()
            findings.append({"rule": f"{m.group(1)}类别应为罗马数字(Ⅰ~Ⅴ)",
                             "loc": loc, "raw": line[max(0, s - 15):s + 20].strip(),
                             "issue": f"出现阿拉伯数字 {m.group(2)}",
                             "expect": "罗马数字"})
        for m in pat_arabic_channel.finditer(line):
            s = m.start()
            findings.append({"rule": "航道/船闸等级应为罗马数字(Ⅰ~Ⅶ)",
                             "loc": loc, "raw": line[max(0, s - 15):s + 20].strip(),
                             "issue": "阿拉伯数字用于航道/船闸等级",
                             "expect": "罗马数字"})
    return findings


# ---------------- 图表编号 gram_007 ----------------
def check_numbering(captions):
    defs = [c for c in captions if c["is_def"]]
    by_type = defaultdict(list)
    for c in defs:
        by_type[c["type"]].append(c)
    result = {}
    for t, items in by_type.items():
        groups = defaultdict(list)
        for c in items:
            mm = re.match(r"^(.+?)[\-\.]([0-9]+)$", c["num"])
            if not mm:
                groups[c["num"]].append((0, c)); continue
            groups[mm.group(1)].append((int(mm.group(2)), c))
        issues = []
        for prefix, lst in groups.items():
            serials = sorted(set(s for s, _ in lst))
            dups = [s for s in serials if sum(1 for x, _ in lst if x == s) > 1]
            gaps = [(serials[i], serials[i + 1])
                    for i in range(len(serials) - 1) if serials[i + 1] - serials[i] > 1]
            if dups:
                issues.append({"prefix": prefix, "type": "重复编号", "detail": f"重复序号 {dups}"})
            if gaps:
                issues.append({"prefix": prefix, "type": "编号间断",
                               "detail": ", ".join(f"{a}->{b}" for a, b in gaps)})
        formats = sorted(set(re.sub(r"\d", "#", c["num"]) for c in items))
        format_inconsistent = len(formats) > 1
        result[t] = {"count": len(items), "formats": formats,
                     "format_inconsistent": format_inconsistent, "issues": issues,
                     "by_prefix": {p: len(v) for p, v in groups.items()}}
    return result


# ---------------- 疑点候选生成（groupC 语义核实的先验，只产候选不产结论） ----------------
DICT_NAME = os.path.join(os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..")), "references", "错字词典.json")

# 合法叠词白名单（正常词语，非错字）
_LEGIT_DOUBLE_WORDS = [
    "往往", "渐渐", "缓缓", "徐徐", "屡屡", "频频", "纷纷", "茫茫", "滔滔", "潺潺",
    "恰恰", "仅仅", "刚刚", "层层", "点点", "丝丝", "阵阵", "熊熊", "绵绵", "皑皑",
    "久久", "深深", "远远", "高高", "长长", "大大", "小小", "多多", "好好", "快快",
    "慢慢", "轻轻", "紧紧", "满满", "圆圆", "匆匆", "一一", "等等", "方方面面",
]
_LEGIT_DOUBLE_RE = re.compile("|".join(_LEGIT_DOUBLE_WORDS))
_DOUBLE_RE = re.compile(r"([\u4e00-\u9fa5])\1")
# 中文语境内的半角标点（前后邻汉字才命中，数字间小数点/千分位不误伤）
_HALF_PUNCT_RE = re.compile(r"(?<=[\u4e00-\u9fa5])[,;:?!](?=[\u4e00-\u9fa5])")
_DUP_PUNCT_RE = re.compile(r"，。|。，|，，|。。|、、|；；|：：")
_MIXED_BRACKET_RE = re.compile(r"（[^（）()]*\)|\([^()（）]*）")
_BRACKET_PAIRS = [("（", "）"), ("【", "】"), ("《", "》"), ("“", "”")]
_MAX_SENT_LEN = 150
_RULE_CAPS = {"叠字": 80, "标点混用": 150, "配对失衡": 60, "长句": 40}
_MAX_HITS_PER_WORD = 5   # 同一词条最多报 5 处（防"砼"这类高频词爆量）


def _load_typo_dict():
    """加载错字词典（references/错字词典.json），只取 必错/偏好/需上下文 三级。"""
    if not os.path.exists(DICT_NAME):
        return []
    try:
        data = json.load(open(DICT_NAME, encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []
    return [e for e in data.get("entries", [])
            if e.get("tier") in ("必错", "偏好", "需上下文") and e.get("wrong")]


def build_suspects(fulltext):
    """生成 groupC 疑点候选：rule + 段落定位 + 上下文 + 提示。
    输出仅是候选，判定一律留给语义层核实，维持『宁可漏报不误报』。"""
    entries = _load_typo_dict()
    dict_by_wrong = {e["wrong"]: e for e in entries}
    word_hits = defaultdict(int)
    caps = defaultdict(int)
    cands = []

    def add(rule, loc, span, hint):
        if caps[rule] >= _RULE_CAPS.get(rule, 10 ** 9):
            return
        caps[rule] += 1
        cands.append({"rule": rule, "loc": loc, "span": span[:60], "hint": hint})

    for line in fulltext.splitlines():
        m = re.match(r"(P\d+)", line)
        if not m:
            continue
        pid = m.group(1)
        text = line[m.end():]
        if len(text) < 2:
            continue

        # 1) 词典（错字/规范偏好/需上下文）
        for w, e in dict_by_wrong.items():
            i = text.find(w)
            while i >= 0 and word_hits[w] < _MAX_HITS_PER_WORD:
                word_hits[w] += 1
                add("词典", pid, text[max(0, i - 15):i + len(w) + 15],
                    "疑为『%s』（%s）" % (e.get("right", ""), e.get("tier", "")))
                i = text.find(w, i + 1)

        # 2) 叠字（先把合法叠词等长占位再扫描）
        masked = _LEGIT_DOUBLE_RE.sub(lambda mm: "〇" * len(mm.group(0)), text)
        for dm in _DOUBLE_RE.finditer(masked):
            ch = dm.group(1)
            if ch == "〇":
                continue
            s = dm.start()
            add("叠字", pid, text[max(0, s - 15):s + 17], "疑为叠字（%s%s）" % (ch, ch))

        # 3) 标点混用
        for pm in _HALF_PUNCT_RE.finditer(text):
            s = pm.start()
            add("标点混用", pid, text[max(0, s - 15):s + 16],
                "中文语境半角标点『%s』" % pm.group(0))
        for pm in _DUP_PUNCT_RE.finditer(text):
            s = pm.start()
            add("标点混用", pid, text[max(0, s - 15):s + 17], "重复标点『%s』" % pm.group(0))
        for pm in _MIXED_BRACKET_RE.finditer(text):
            s = pm.start()
            add("标点混用", pid, text[max(0, s - 15):s + 17], "全半角括号混用")

        # 4) 括号/引号行级配对
        for l, r in _BRACKET_PAIRS:
            if text.count(l) != text.count(r):
                add("配对失衡", pid, text[:50],
                    "『%s%s』数量不等（%d/%d）" % (l, r, text.count(l), text.count(r)))
                break

        # 5) 超长句
        for sent in re.split(r"[。；！？]", text):
            if len(sent) > _MAX_SENT_LEN:
                add("长句", pid, sent[:60] + "…（%d 字无句读）" % len(sent), "超长句，通顺性待核")

    summary = {"total": len(cands)}
    for c in cands:
        summary[c["rule"]] = summary.get(c["rule"], 0) + 1
    return {"summary": summary, "candidates": cands}


def main():
    args = sys.argv[1:]
    run_dir = args[0] if len(args) > 0 else None
    here = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.normpath(os.path.join(here, ".."))
    if not run_dir:
        rd = os.path.join(skill_root, "_run")
        subs = [os.path.join(rd, x) for x in os.listdir(rd)] if os.path.isdir(rd) else []
        subs = [s for s in subs if os.path.isdir(s)]
        run_dir = max(subs, key=lambda p: os.path.getmtime(p)) if subs else None
    if not run_dir or not os.path.isdir(run_dir):
        print("用法: python deterministic_checks.py <run_dir>"); sys.exit(1)

    global INDEX
    INDEX = json.load(open(os.path.join(here, "references_index.json"), encoding="utf-8"))

    tables = json.load(open(os.path.join(run_dir, "tables.json"), encoding="utf-8"))
    captions = json.load(open(os.path.join(run_dir, "captions.json"), encoding="utf-8"))
    citations = json.load(open(os.path.join(run_dir, "citations.json"), encoding="utf-8"))
    numbers = json.load(open(os.path.join(run_dir, "numbers.json"), encoding="utf-8"))
    fulltext = open(os.path.join(run_dir, "fulltext.md"), encoding="utf-8").read()

    det_table = check_tables(tables)
    det_cit = check_citations(citations, INDEX)
    det_num = check_numbers(numbers, fulltext)
    det_num7 = check_numbering(captions)
    det_sus = build_suspects(fulltext)

    for name, data in (("det_table.json", det_table), ("det_citations.json", det_cit),
                       ("det_numbers.json", det_num), ("det_numbering.json", det_num7),
                       ("det_suspects.json", det_sus)):
        json.dump(data, open(os.path.join(run_dir, name), "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)

    print("[deterministic_checks] 摘要：")
    print(f"  表格核算：涉及 {len(det_table)} 张表，"
          f"不合格 {sum(t['n_fail'] for t in det_table)} 项，存疑 {sum(t['n_warn'] for t in det_table)} 项")
    print(f"  引用：{det_cit['summary']}")
    print(f"  引用一致性问题：{det_cit['summary']['consistency_issues']} 组")
    print(f"  流量有效数字(std_003) 违规 {det_num['std003_flow']['n_violations']} 处")
    print(f"  大数单位(std_001) 文字单位={det_num['std001_bigunit']['word_forms']} "
          f"10的次方={det_num['std001_bigunit']['pow10_count']} 混用={det_num['std001_bigunit']['mixed_form']}")
    print(f"  等别/级别(std_004) 异常 {sum(1 for g in det_num['std004_grade'] if not g.get('ok'))} 处")
    for t, r in det_num7.items():
        print(f"  图表编号({t}) 共 {r['count']} 个题录，格式 {r['formats']}，"
              f"格式不一致={r['format_inconsistent']}，序号问题 {len(r['issues'])} 组")
    print(f"  疑点候选(C组先验) 共 {det_sus['summary']['total']} 条：{det_sus['summary']}")


if __name__ == "__main__":
    main()
