# -*- coding: utf-8 -*-
"""round_diff.py — 「轮次增量模式」判定器（只读，不修改任何输入）

用途
----
判定某个 docx 润色项目的第 N 轮（N>=2）相对上一轮**到底变了哪些段**，并输出后续步骤
可直接消费的结构：changed_ids / 变化类型 / 保护清单提示 / 建议分片 / 全篇级决策观察点 /
token 收益提示。它是「增量模式」的前置闸门：只有它算出的 changed_ids 才允许进改写遍。

对齐策略（关键）
----------------
历史语料实测显示：**id 与 paraIndex 都会漂移**（编制人删/插段落会让其后所有 id 整体平移，
例如信江 r2 相对 r1 少了 2 段 ⇒ 按 id 对齐只能命中 192/2197，按文本锚定对齐命中 2181/2197）。
因此本脚本默认 `--align auto`：先试 id 序，失败则改用**文本序列锚定对齐**
（difflib.SequenceMatcher，对 equal/insert/delete/replace 逐块处理，replace 块再按
位置+模糊相似度配对），保证「删了一段」不会把后面所有段误判为"已修改"。

输出（--out 写 JSON；结构见文件末 EPILOG）
-----------------------------------------
counts / chars / changed_ids / changed[] / added_ids / removed_ids /
protection / global_watch / suggested_shards / token_hint / warnings

用法
----
  # 最常用：给新稿 + 上轮工作目录，自动挑上轮最终产物
  python round_diff.py --new D:\\..\\lushui3_polish\\paragraphs.json \\
                       --prev-dir D:\\..\\lushui2_polish --out diff.json --md diff.md

  # 显式指定两文件，并载入上轮保护清单做提示
  python round_diff.py --new new\\paragraphs.json --prev old\\final.json \\
                       --protect old\\protected_ids.json --out diff.json

  # 逐段元数据（含 omml / run 级上下标）与源 docx 的 numPr 检测
  python round_diff.py --new ... --prev ... --meta old\\fmt_scan.json --docx src.docx

  # 只打印摘要，不落盘
  python round_diff.py --new ... --prev ... --json

约定
----
* 纯只读：只读入参文件，只在 --out/--md/--csv 指定路径写文件。
* UTF-8 无 BOM 输出；stdout 强制 UTF-8。
* 退出码：0 正常；2 参数/文件错误；3 发现「上轮产物不可判定」等需要人工确认的情况。
"""
from __future__ import annotations

import argparse
import csv
import difflib
import io
import json
import os
import re
import sys

SCHEMA = "round_diff/1"

# ---------------------------------------------------------------- 基础 IO

def _fix_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def load_json(path):
    with io.open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def load_paras(path):
    """返回 [(id, paraIndex, text)]；兼容 paragraphs.json 与 merged*/final*.json"""
    d = load_json(path)
    if isinstance(d, dict):
        for k in ("paragraphs", "items", "data", "result"):
            if isinstance(d.get(k), list):
                d = d[k]
                break
    if not isinstance(d, list):
        raise ValueError("不是 JSON 数组: %s" % path)
    out = []
    for i, x in enumerate(d):
        if not isinstance(x, dict):
            raise ValueError("元素不是对象: %s[%d]" % (path, i))
        t = x.get("text")
        if t is None:
            t = x.get("newText")
        if t is None:
            t = x.get("original")
        out.append((x.get("id"), x.get("paraIndex", i), t if t is not None else ""))
    return out


# ---------------------------------------------------------------- 对齐

def _norm(s, ignore_space):
    if not ignore_space:
        return s
    return re.sub(r"\s+", "", s or "")


def align(prev, new, fuzzy=0.60, ignore_space=False, fuzzy_max_block=120, mode="auto"):
    """锚定对齐（多通道，先精确后模糊，避免 O(n^2)）。

    mode="auto"（默认）：按下面五通道顺序对齐，changed_ids 可信。
    mode="id"  ：**朴素口径**，只按 id 对齐（同 id 文本不同即算"已改"）。这是很多脚本的默认做法，
                 一旦 id/顺序漂移就会把漂移误判成改动（信江 r2 实测 16 段真变化被报成 2005 段），
                 只用于复现对照，不要用它做增量判定。
    mode="text"：跳过 id/index 通道，只用文本锚定（用于验证 id 完全不可信时的结果）。

    通道顺序（命中即记入 channels，供诊断"id 漂移有多严重"）：
      id      —— 同 id 且文本相同
      index   —— 同 paraIndex 且文本相同（id 编号体系变了但顺序没变）
      text    —— 文本相同、位置漂移（= 有段落插入/删除把后面整体平移）
      order   —— 残余段用 difflib 顺序对齐判 equal
      fuzzy   —— 残余 replace 块内相似度 >= fuzzy 的配对 = 已修改段
    剩下未配 = added（本轮新增）/ removed（上轮独有）。

    返回 dict(unchanged=[(j_new, i_prev)], modified=[(j_new,i_prev,sim)],
              added=[j_new], removed=[i_prev], channels={...}, id_order_equal=bool)
    """
    import bisect
    pt = [_norm(x[2], ignore_space) for x in prev]
    nt = [_norm(x[2], ignore_space) for x in new]
    n_p, n_n = len(pt), len(nt)
    channels = {"id": 0, "index": 0, "text": 0, "order": 0, "fuzzy": 0}
    id_order_equal = [x[0] for x in prev] == [x[0] for x in new]

    if mode == "id":
        pmap = {}
        for i, x in enumerate(prev):
            pmap.setdefault(x[0], i)
        unchanged, modified, added, removed = [], [], [], []
        used = set()
        for j, x in enumerate(new):
            i = pmap.get(x[0])
            if i is None:
                added.append(j)
                continue
            used.add(i)
            if pt[i] == nt[j]:
                unchanged.append((j, i))
                channels["id"] += 1
            else:
                r = difflib.SequenceMatcher(None, pt[i], nt[j]).ratio()
                modified.append((j, i, r))
                channels["fuzzy"] += 1
        removed = [i for i in range(n_p) if i not in used]
        modified.sort()
        return dict(unchanged=unchanged, modified=modified, added=added, removed=removed,
                    channels=channels, id_order_equal=id_order_equal)

    pu = [False] * n_p
    nu = [False] * n_n
    unchanged, modified = [], []

    def take(j, i, ch):
        pu[i] = True
        nu[j] = True
        unchanged.append((j, i))
        channels[ch] = channels.get(ch, 0) + 1

    use_id = mode in ("auto", "id")
    # 通道 1：同 id 且同文本
    if use_id:
        by_id = {}
        for i, x in enumerate(prev):
            by_id.setdefault(x[0], []).append(i)
        for j, x in enumerate(new):
            for i in by_id.get(x[0], ()):
                if not pu[i] and pt[i] == nt[j]:
                    take(j, i, "id")
                    break
        # 通道 2：同 paraIndex 且同文本
        by_idx = {}
        for i, x in enumerate(prev):
            by_idx.setdefault(x[1], []).append(i)
        for j, x in enumerate(new):
            if nu[j]:
                continue
            for i in by_idx.get(x[1], ()):
                if not pu[i] and pt[i] == nt[j]:
                    take(j, i, "index")
                    break
    # 通道 3：文本相同、位置漂移（取离 j 最近的上轮下标）。
    #          注意：重复短段（"小计"/"附表"等）在此通道可能被交叉配对，这不影响
    #          changed_ids（把它们正确判成"未变"就够），只影响漂移分段的报告；
    #          drift_segments 内部用最长递增子序列过滤掉这类孤立交叉配对。
    pos = {}
    for i, t in enumerate(pt):
        if not pu[i]:
            pos.setdefault(t, []).append(i)
    for j in range(n_n):
        if nu[j]:
            continue
        cand = pos.get(nt[j])
        if not cand:
            continue
        if len(cand) > 64:          # 高频重复段（如空段/固定表头）：取第一个未用即可
            pick = None
            for i in cand:
                if not pu[i]:
                    pick = i
                    break
            if pick is None:
                continue
        else:
            k = bisect.bisect_left(cand, min(j, n_p - 1))
            pick = None
            for d in range(len(cand)):
                for t2 in (k - d, k + d):
                    if 0 <= t2 < len(cand) and not pu[cand[t2]]:
                        pick = cand[t2]
                        break
                if pick is not None:
                    break
            if pick is None:
                continue
        take(j, pick, "text")

    # 通道 4/5：残余序列做顺序对齐 + 块内模糊配对
    res_p = [i for i in range(n_p) if not pu[i]]
    res_n = [j for j in range(n_n) if not nu[j]]
    if res_p or res_n:
        rpt = [pt[i] for i in res_p]
        rnt = [nt[j] for j in res_n]
        sm = difflib.SequenceMatcher(None, rpt, rnt, autojunk=False)
        added, removed = [], []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                for k in range(i2 - i1):
                    take(res_n[j1 + k], res_p[i1 + k], "order")
            elif tag == "delete":
                removed.extend(res_p[i1:i2])
            elif tag == "insert":
                added.extend(res_n[j1:j2])
            else:
                a = res_p[i1:i2]
                b = res_n[j1:j2]
                if len(a) * len(b) <= fuzzy_max_block * fuzzy_max_block:
                    pairs = []
                    for j in b:
                        best = None
                        for i in a:
                            r = difflib.SequenceMatcher(None, pt[i], nt[j]).ratio()
                            if best is None or r > best[0]:
                                best = (r, i)
                        if best and best[0] >= fuzzy:
                            pairs.append((best[0], best[1], j))
                    pairs.sort(reverse=True)
                    ua, ub = set(), set()
                    for r, i, j in pairs:
                        if i in ua or j in ub:
                            continue
                        ua.add(i)
                        ub.add(j)
                        pu[i] = True
                        nu[j] = True
                        modified.append((j, i, r))
                        channels["fuzzy"] += 1
                    for i in a:
                        if i not in ua:
                            removed.append(i)
                    for j in b:
                        if j not in ub:
                            added.append(j)
                else:  # 大块不回退到 O(n^2)：按位置配对，相似度现算
                    for i, j in zip(a, b):
                        r = difflib.SequenceMatcher(None, pt[i], nt[j]).ratio()
                        if r >= fuzzy:
                            pu[i] = True
                            nu[j] = True
                            modified.append((j, i, r))
                            channels["fuzzy"] += 1
                        else:
                            removed.append(i)
                            added.append(j)
                    for i in a[len(b):]:
                        removed.append(i)
                    for j in b[len(a):]:
                        added.append(j)
    else:
        added, removed = [], []

    modified.sort()
    added.sort()
    removed.sort()
    id_order_equal = [x[0] for x in prev] == [x[0] for x in new]
    return dict(unchanged=unchanged, modified=modified, added=added,
                removed=removed, channels=channels, id_order_equal=id_order_equal)


def _lis_subset(pairs):
    """对已按 new 序排列的 (j,i) 配对取「最长递增（按 i）」子序列。

    用途：重复短段会被文本通道交叉配对，造成偏移量在局部来回跳；取 LIS 后剩下的
    配对在顺序上自洽，偏移量分段才代表真实的段落增删位置。（O(n log n)）
    """
    import bisect
    pairs = list(pairs)
    tails = []      # tails[k] = 最小的 i
    tail_k = []     # tails[k] 对应的 pairs 下标
    back = [-1] * len(pairs)
    for k, (j, i) in enumerate(pairs):
        p = bisect.bisect_left(tails, i)
        if p == len(tails):
            tails.append(i)
            tail_k.append(k)
        else:
            tails[p] = i
            tail_k[p] = k
        back[k] = tail_k[p - 1] if p > 0 else -1
    if not tail_k:
        return []
    out = []
    k = tail_k[-1]
    while k >= 0:
        out.append(pairs[k])
        k = back[k]
    out.reverse()
    return out


def drift_segments(prev, new, unchanged, min_len=2):
    """顺序漂移分段：用「对齐上的未变段」计算 offset = 上轮序号 - 本轮序号。

    offset 在一个区间内保持不变 = 该区间内两轮段落一一对应；
    offset 跳变处 = 发生了段落增删（其后所有段的 paraIndex 整体平移）。
    先用 LIS 过滤掉重复段交叉配对造成的孤立谬配对，再按 offset 分段。
    返回 [(new_para_index_start, new_para_index_end, offset, count)]
    """
    pairs = _lis_subset(sorted(unchanged))
    segs = []
    cur = None
    for j, i in pairs:
        off = i - j
        if cur is None:
            cur = [j, j, off, 1]
        elif off == cur[2]:
            cur[1] = j
            cur[3] += 1
        else:
            segs.append(cur)
            cur = [j, j, off, 1]
    if cur:
        segs.append(cur)
    out = []
    for a, b, off, cnt in segs:
        if cnt >= min_len:
            out.append(dict(new_para_index_start=a, new_para_index_end=b,
                            id_start=new[a][0] if a < len(new) else None,
                            id_end=new[b][0] if b < len(new) else None,
                            offset=off, count=cnt))
    return out


# ---------------------------------------------------------------- 保护清单 / 全篇级观察点

RE_REF = re.compile(r"^\s*\[\d+\]")
RE_STD_IN_LINE = re.compile(r"(GB|SL|DL|DB|JT|QX|CJJ|JTG|HJ|NY|T/CEC|SL/T)\s*/?\s*T?\s*\d{2,}")
RE_CONFIG = re.compile(r"^\s*\[[A-Za-z_][A-Za-z0-9_]*\]|^\s*//|\b(intVal|doubleVal|strVal|floatVal)\b")
RE_ASSIGN = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\s*=\s*[-+]?[\d.]")
RE_SYMBOL = re.compile(r"^[\x00-\x7F\s]*$")
RE_CAPTION = re.compile(r"^\s*(图|表|附图|附表)\s*[0-9一二三四五六七八九十]")
RE_TOC = re.compile(r"[.·]{3,}\s*\d+\s*$")
RE_HEADING = re.compile(r"^\s*(\d+(\.\d+)*\s|第[一二三四五六七八九十百]+[章节篇]|[一二三四五六七八九十]+、)")
RE_LISTONLY = re.compile(r"^\s*[（(]?\s*[0-9一二三四五六七八九十]+\s*[）).、]?\s*$")
RE_SUPSUB_UNICODE = re.compile(r"[²³¹⁰⁴⁵⁶⁷⁸⁹₀-₉]")
RE_POWERTEXT = re.compile(r"(km|cm|mm|dm|m|hm|kg|t|L|s|h)[23]\b", re.I)
RE_OMML_HINT = re.compile(r"[=∑∏∫√±×÷≈≤≥∂∆∇αβγδεζηθλμπρστφχψω]")

G_UNIT = re.compile(r"m³|m3|km²|km2|m²|m2|hm²|L/s|m/s|万\s*m³|亿\s*m³|10\s*[\^*]?\s*[³3]")
G_UNIT2 = re.compile(r"\d\s*(km|m|mm|cm|hm|kg|t|万|亿|亩|人|万元|m³|km²)(?![a-zA-Z])")
G_FIGTAB = re.compile(r"(图|表|附图|附表)\s*\d")
G_STD = re.compile(r"(GB|SL|DL|DB|JT|QX|CJJ|JTG|HJ|NY)\s*/?\s*T?\s*\d")
G_LAW = re.compile(r"第[一二三四五六七八九十百零〇\d]{1,8}号|〔\d{4}〕|\[\d{4}\]")
G_QUOTE = re.compile(r"《|》")


def protect_reasons(text, meta=None):
    """返回该段命中的保护/观察规则列表 [(rule, reason)]"""
    hits = []
    t = text or ""
    if RE_REF.search(t) and (RE_STD_IN_LINE.search(t) or re.search(r"\[[A-Z]\]|\[D\]|\[J\]|\[M\]", t)):
        hits.append(("REF_LIST", "参考文献/编制依据清单行：GB/T 7714 著录体例，整行原样"))
    if RE_REF.search(t):
        hits.append(("REF_LINE", "以 [n] 开头的条目行：清单体例，整行原样"))
    if RE_CONFIG.search(t) or RE_ASSIGN.search(t):
        hits.append(("CONFIG_LINE", "软件输入文件/参数说明行：模型原文示例，整行原样"))
    if t and RE_SYMBOL.match(t) and len(t.strip()) > 2 and re.search(r"[A-Za-z]", t):
        hits.append(("ASCII_LINE", "纯 ASCII 行（英文/符号/变量）：不属中文润色对象"))
    if RE_CAPTION.match(t):
        hits.append(("CAPTION", "图题/表题：编号与文字均为全篇级对象，须由父代理统一"))
    if RE_TOC.match(t):
        hits.append(("TOC", "目录行（含页码点线）：不得改写"))
    if RE_LISTONLY.match(t):
        hits.append(("LIST_ONLY", "纯序号行：序号可能由 Word 自动生成（numPr），写文本会双序号"))
    if RE_SUPSUB_UNICODE.search(t) or RE_POWERTEXT.search(t):
        hits.append(("SUPSUB", "含上/下标字符或单位幂次：须核对 run 级 vertAlign，禁止增删幂次"))
    if RE_OMML_HINT.search(t):
        hits.append(("FORMULA_HINT", "疑似公式/含数学符号：整段原样（OMML 重建会丢公式）"))
    if len(t.strip()) <= 4:
        hits.append(("SHORT", "极短段（<=4 字）：改动比会失真，禁止自动改写"))
    if meta:
        if meta.get("omml"):
            hits.append(("OMML", "元数据显示含 OMML 公式（omml=%s）：必须整段保护" % meta["omml"]))
        if meta.get("sups"):
            hits.append(("RUN_SUP", "元数据显示含 run 级上标（sups=%s）：文本层是普通数字，禁止转 Unicode"
                         % json.dumps(meta["sups"], ensure_ascii=False)[:40]))
        if meta.get("bookmarks"):
            hits.append(("BOOKMARK", "含书签（含 _Toc 锚点）：重建段落会丢锚点"))
        st = (meta.get("style") or "").strip()
        if st:
            hits.append(("STYLE", "非正文样式 style=%s：标题/题注建议只做全篇级统一" % st))
    if meta and meta.get("numpr"):
        hits.append(("NUMPR", "源 docx 该段带 numPr 自动编号（numId=%s）：写入文本会出现双序号，必须回退"
                     % meta["numpr"]))
    return hits


def global_watch_reasons(text):
    hits = []
    t = text or ""
    if G_UNIT.search(t) or G_UNIT2.search(t):
        hits.append(("UNIT_STYLE", "含单位/量纲写法：m³/m3、上标下标属全篇体例决策"))
    if G_FIGTAB.search(t):
        hits.append(("FIGTAB_REF", "含图表编号或引用：编号连续性与引用一致性属全篇决策"))
    if G_STD.search(t) or RE_STD_IN_LINE.search(t):
        hits.append(("STANDARD_ID", "含法规/标准编号：须与全篇标准表交叉核对"))
    if G_LAW.search(t):
        hits.append(("LAW_ID", "含文号/令号/年份括号：法规引用属全篇决策"))
    if G_QUOTE.search(t):
        hits.append(("BOOKTITLE", "含《》书名号：标准/文件名须全篇统一"))
    return hits


# ---------------------------------------------------------------- 源 docx 元数据（可选）

def scan_docx(path):
    """读源 docx（只读），返回 {paraIndex: {...}} 含 numPr / omml / run 级上下标 / 书签。

    复用既有案例的做法：正文段 = body 直属 w:p，序号即 paraIndex。
    """
    from docx import Document  # 延迟导入
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    doc = Document(path)
    paras = [c for c in doc.element.body if c.tag == W + "p"]
    out = {}
    for idx, p in enumerate(paras):
        rec = {}
        ppr = p.find(W + "pPr")
        if ppr is not None:
            numpr = ppr.find(W + "numPr")
            if numpr is not None:
                nid = numpr.find(W + "numId")
                ilvl = numpr.find(W + "ilvl")
                rec["numpr"] = "%s/%s" % (
                    nid.get(W + "val") if nid is not None else "?",
                    ilvl.get(W + "val") if ilvl is not None else "?")
            ps = ppr.find(W + "pStyle")
            if ps is not None:
                rec["style"] = ps.get(W + "val") or ""
        if p.findall(".//" + M + "oMath") or p.findall(".//" + M + "oMathPara"):
            rec["omml"] = len(p.findall(".//" + M + "oMath"))
        sups = []
        for r in p.findall(".//" + W + "r"):
            rpr = r.find(W + "rPr")
            if rpr is None:
                continue
            va = rpr.find(W + "vertAlign")
            if va is not None and va.get(W + "val") in ("superscript", "subscript"):
                t = "".join(x.text or "" for x in r.findall(W + "t"))
                sups.append(t)
        if sups:
            rec["sups"] = sups
        bms = [b.get(W + "name") for b in p.findall(".//" + W + "bookmarkStart")]
        bms = [b for b in bms if b]
        if bms:
            rec["bookmarks"] = bms
        out[idx] = rec
    return out


def load_meta(path):
    """fmt_scan.json 风格：list[{id, omml, sups, bookmarks, style}]"""
    d = load_json(path)
    by_id = {}
    if isinstance(d, list):
        for x in d:
            if isinstance(x, dict) and "id" in x:
                by_id[x["id"]] = x
    return by_id


def load_protect(paths):
    """支持 {id:reason} / [id,...] / {"ids":[...]} / [{"id":..,"reason":..}]"""
    ids, reasons = set(), {}
    for p in paths:
        d = load_json(p)
        if isinstance(d, dict):
            if isinstance(d.get("ids"), list):
                for i in d["ids"]:
                    ids.add(_as_id(i))
                continue
            for k, v in d.items():
                i = _as_id(k)
                ids.add(i)
                if isinstance(v, str):
                    reasons[i] = v
        elif isinstance(d, list):
            for x in d:
                if isinstance(x, dict) and "id" in x:
                    i = _as_id(x["id"])
                    ids.add(i)
                    if x.get("reason"):
                        reasons[i] = str(x["reason"])
                else:
                    ids.add(_as_id(x))
    return ids, reasons


def _as_id(x):
    if isinstance(x, str) and x.strip().lstrip("-").isdigit():
        return int(x)
    return x


# ---------------------------------------------------------------- 主流程

PREV_PRIORITY = ["merged_final.json", "final.json", "merged.json"]


def pick_prev(dirpath):
    for name in PREV_PRIORITY:
        p = os.path.join(dirpath, name)
        if os.path.exists(p):
            return p, name
    return None, None


def build(args):
    new = load_paras(args.new)
    prev = load_paras(args.prev)
    al = align(prev, new, fuzzy=args.fuzzy, ignore_space=args.ignore_space,
               mode=getattr(args, "align", "auto"))

    meta_by_id = load_meta(args.meta) if args.meta else {}
    dx = {}
    if args.docx:
        dx = scan_docx(args.docx)
    pids, preasons = load_protect(args.protect or [])

    total_chars = sum(len(x[2] or "") for x in new)
    changed = []
    changed_ids = []
    changed_chars = 0
    for j, i, sim in al["modified"]:
        t = new[j][2] or ""
        changed.append(dict(id=new[j][0], prev_id=prev[i][0], para_index=new[j][1],
                            align="text", sim=round(sim, 4),
                            change_type="edited" if sim >= 0.60 else "rewritten",
                            len_new=len(t), len_prev=len(prev[i][2] or ""), added=False))
        changed_ids.append(new[j][0])
        changed_chars += len(t)
    for j in al["added"]:
        t = new[j][2] or ""
        changed.append(dict(id=new[j][0], prev_id=None, para_index=new[j][1],
                            align="none", sim=0.0, change_type="added",
                            len_new=len(t), len_prev=0, added=True))
        changed_ids.append(new[j][0])
        changed_chars += len(t)
    changed.sort(key=lambda x: (x["para_index"] if isinstance(x["para_index"], int) else 0))

    unchanged_ids = [new[j][0] for j, _ in al["unchanged"]]

    # id 漂移
    drift = drift_segments(prev, new, al["unchanged"])
    warnings = []
    if not al["id_order_equal"]:
        warnings.append("id 序列在上轮与本轮之间不一致（id 漂移）：本脚本已改用文本锚定对齐；"
                        "任何按 id 直连的比较结果都会被高估，禁止用 id 交集当 changed_ids")
    if al["added"] or al["removed"]:
        warnings.append("存在段落级增删：新增 %d 段 / 上轮独有 %d 段；"
                        "其后所有段的 id/paraIndex 均可能平移" % (len(al["added"]), len(al["removed"])))
    if len(prev) != len(new):
        warnings.append("段数不等（上轮 %d / 本轮 %d）" % (len(prev), len(new)))
    overlap = len(al["unchanged"]) / float(len(new) or 1)
    if overlap < args.min_overlap:
        warnings.append(
            "本轮输入与上轮产物的文本重合率只有 %.4f（<%s）：**很可能不是相邻轮次**"
            "（跳轮/上轮产物选错/换了稿）。实测相邻轮次为 0.9733~0.9958；"
            "同模板的兄弟分册之间可达 0.5585~0.6504，跳一轮（r1→r3）为 0.8344。"
            "此时 changed_ids 只能当'相对该基线'的差异，不能当'本轮人工改动'"
            % (overlap, args.min_overlap))

    # 保护与全篇级观察
    def reasons_for(j):
        """返回该新稿下标 j 命中的全部保护规则（启发式 + 元数据 + docx numPr/OMML/上下标 + 声明清单）"""
        x = new[j]
        reasons = list(protect_reasons(x[2], meta_by_id.get(x[0])))
        if dx:
            m = dx.get(x[1])
            if m:
                reasons.extend(protect_reasons(x[2], m))
        if x[0] in pids:
            reasons.append(("DECLARED", "在上轮保护清单中：%s" % preasons.get(x[0], "")))
        return reasons

    def enrich(rec, idx_in_new):
        reasons = reasons_for(idx_in_new)
        rec["protected"] = bool(reasons)
        rec["protect_rules"] = sorted({r for r, _ in reasons})
        rec["protect_reasons"] = [{"rule": r, "reason": d} for r, d in reasons]
        return rec

    idx_by_id = {}
    for j, x in enumerate(new):
        idx_by_id.setdefault(x[0], j)
    for rec in changed:
        j = idx_by_id.get(rec["id"])
        if j is not None:
            enrich(rec, j)

    cidset = set(changed_ids)
    prot_hits = []
    rule_count = {}
    for j, x in enumerate(new):
        reasons = reasons_for(j)
        for r, _ in reasons:
            rule_count[r] = rule_count.get(r, 0) + 1
        if reasons:
            prot_hits.append(dict(id=x[0], para_index=x[1], rules=sorted({r for r, _ in reasons}),
                                  in_changed=(x[0] in cidset),
                                  declared=(x[0] in pids)))
    watch_hits = []
    wrule = {}
    for x in new:
        rs = global_watch_reasons(x[2])
        for r, _ in rs:
            wrule[r] = wrule.get(r, 0) + 1
        if rs:
            watch_hits.append(dict(id=x[0], para_index=x[1], rules=sorted({r for r, _ in rs}),
                                   in_changed=(x[0] in cidset)))

    # 「按 id 直连」的朴素口径（任务书要求如实报）：仅供漂移对照，不作结论依据
    prev_by_id = {}
    for i, x in enumerate(prev):
        prev_by_id.setdefault(x[0], x[2])
    naive_same_id = sum(1 for x in new if prev_by_id.get(x[0]) == x[2])
    naive_cmp_id = sum(1 for x in new if x[0] in prev_by_id)
    naive_same_idx = sum(1 for j, x in enumerate(new) if j < len(prev) and prev[j][2] == x[2])

    # 建议分片：按文档序累积字数切分（只含 changed_ids）
    shards = []
    if changed_ids:
        cid = set(changed_ids)
        cur_ids, cur_chars = [], 0
        for x in new:
            if x[0] not in cid:
                continue
            cur_ids.append(x[0])
            cur_chars += len(x[2] or "")
            if cur_chars >= args.shard_chars:
                shards.append(dict(name="shard%02d" % (len(shards) + 1), count=len(cur_ids),
                                   chars=cur_chars, ids=cur_ids))
                cur_ids, cur_chars = [], 0
        if cur_ids:
            shards.append(dict(name="shard%02d" % (len(shards) + 1), count=len(cur_ids),
                               chars=cur_chars, ids=cur_ids))

    ratio = (changed_chars / total_chars) if total_chars else 0.0
    result = dict(
        schema=SCHEMA,
        new=dict(path=os.path.abspath(args.new), count=len(new), chars=total_chars),
        prev=dict(path=os.path.abspath(args.prev), count=len(prev),
                  chars=sum(len(x[2] or "") for x in prev)),
        alignment=dict(mode=getattr(args, "align", "auto"), id_order_equal=al["id_order_equal"],
                       channels=al["channels"],
                       id_hit_ratio=round(al["channels"].get("id", 0) / max(len(new), 1), 4),
                       text_overlap_ratio=round(overlap, 4),
                       drift_segments=drift,
                       note="changed_ids 来自锚定对齐（id/index/text/order/fuzzy 五通道）；"
                            "id_hit_ratio 越低说明 id 漂移越严重，按 id 直连的比较结果越不可用"),
        counts=dict(paras=len(new), paras_prev=len(prev),
                    unchanged=len(al["unchanged"]), modified=len(al["modified"]),
                    added=len(al["added"]), removed=len(al["removed"]),
                    changed_ids=len(changed_ids)),
        chars=dict(total=total_chars, changed=changed_chars,
                   changed_ratio=round(ratio, 6),
                   unchanged=total_chars - changed_chars),
        naive=dict(
            note="朴素口径（直接按 id/序号同位置比对）——**只作漂移证据，不作为 changed_ids**",
            by_id_compared=naive_cmp_id,
            by_id_same=naive_same_id,
            by_id_diff=naive_cmp_id - naive_same_id,
            by_index_same=naive_same_idx,
            by_index_diff=len(new) - naive_same_idx),
        changed_ids=changed_ids,
        changed=changed,
        unchanged_ids=unchanged_ids,
        added_ids=[new[j][0] for j in al["added"]],
        removed_ids=[prev[i][0] for i in al["removed"]],
        protection=dict(declared=len(pids), hits=prot_hits, by_rule=rule_count,
                        hits_in_changed=sum(1 for h in prot_hits if h["in_changed"])),
        global_watch=dict(hits=watch_hits, by_rule=wrule,
                          note="这些段承载全篇级决策（单位体例/上下标/图表编号/法规标准编号）；"
                               "即使未变，也必须全量重扫，不能只看 changed_ids"),
        suggested_shards=shards,
        token_hint=dict(
            changed_chars=changed_chars, total_chars=total_chars,
            output_token_saving_ratio=round(ratio, 6),
            note="改写遍的输出 token 与输入段落文本近似线性；output_token_saving_ratio=1-变化段字数占比。"
                 "分片固定成本（任务书 prompt + lint 清单）不随段数缩小，见 report.md 的估算式"),
        warnings=warnings,
    )
    return result


# ---------------------------------------------------------------- 输出

def summary_md(res, prev_pick=None):
    c = res["counts"]
    ch = res["chars"]
    L = []
    L.append("| 指标 | 值 |")
    L.append("|---|---|")
    L.append("| 新稿 | `%s`（%d 段 / %d 字） |" % (res["new"]["path"], res["new"]["count"], res["new"]["chars"]))
    L.append("| 上轮产物 | `%s`（%d 段 / %d 字） |" % (res["prev"]["path"], res["prev"]["count"], res["prev"]["chars"]))
    L.append("| 未变 / 已改 / 新增 / 上轮独有 | %d / %d / %d / %d |"
             % (c["unchanged"], c["modified"], c["added"], c["removed"]))
    L.append("| changed_ids | %d |" % c["changed_ids"])
    L.append("| 变化段字数占比 | %.4f%%（%d / %d 字） |" % (100 * ch["changed_ratio"], ch["changed"], ch["total"]))
    L.append("| id 序是否一致 | %s |" % ("是" if res["alignment"]["id_order_equal"] else "**否（漂移）**"))
    L.append("| 与上轮文本重合率 | %.4f %s |"
             % (res["alignment"]["text_overlap_ratio"],
                "（≥0.95，符合相邻轮次）" if res["alignment"]["text_overlap_ratio"] >= 0.95
                else "**（<0.95，疑非相邻轮次）**"))
    L.append("| id 通道命中率 | %.2f%%（id/index/text/order/fuzzy = %s） |"
             % (100 * res["alignment"]["id_hit_ratio"],
                "/".join(str(res["alignment"]["channels"].get(k, 0))
                         for k in ("id", "index", "text", "order", "fuzzy"))))
    nv = res.get("naive", {})
    if nv:
        L.append("| 朴素口径（同 id 比） | 比 %d 段 / 相同 %d / 不同 %d ← **按 id 直连会把漂移误判成改动** |"
                 % (nv["by_id_compared"], nv["by_id_same"], nv["by_id_diff"]))
        L.append("| 朴素口径（同序号比） | 相同 %d / 不同 %d |"
                 % (nv["by_index_same"], nv["by_index_diff"]))
    L.append("| 保护清单声明 / 命中 / 命中且已变 | %d / %d / %d |"
             % (res["protection"]["declared"], len(res["protection"]["hits"]),
                res["protection"]["hits_in_changed"]))
    L.append("| 全篇级观察点段数 | %d |" % len(res["global_watch"]["hits"]))
    L.append("| 建议分片 | %d 片（%s） |" % (len(res["suggested_shards"]),
              ", ".join("%s:%d段" % (s["name"], s["count"]) for s in res["suggested_shards"]) or "-"))
    if res["warnings"]:
        L.append("")
        L.append("**告警**")
        for w in res["warnings"]:
            L.append("- %s" % w)
    if res["alignment"]["drift_segments"]:
        L.append("")
        L.append("**顺序漂移分段（offset = 上轮序号 - 本轮序号；offset 跳变处 = 发生段落增删）**")
        L.append("")
        L.append("| 本轮 paraIndex 起 | 止 | 本轮 id 起 | offset | 段数 |")
        L.append("|---|---|---|---|---|")
        for s in res["alignment"]["drift_segments"][:40]:
            L.append("| %s | %s | %s | %s | %d |"
                     % (s["new_para_index_start"], s["new_para_index_end"],
                        s["id_start"], s["offset"], s["count"]))
    return "\n".join(L)


def main(argv=None):
    _fix_stdout()
    ap = argparse.ArgumentParser(
        prog="round_diff.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="轮次增量模式判定器：算出第 N 轮相对上一轮真正变化的段，输出 changed_ids/"
                    "变化类型/保护清单提示/建议分片/全篇级观察点（只读）",
        epilog="""输出 JSON 字段（后续步骤消费点）
  changed_ids[]         —— 唯一允许进改写遍的段 id（文本锚定对齐得出）
  changed[]             —— {id, prev_id, para_index, sim, change_type, len_new/len_prev,
                            protected, protect_rules, protect_reasons}
                            change_type: edited(0.6<=sim<1) / rewritten(sim<0.6) / added
  unchanged_ids[]       —— 未变段（增量模式不送改写；但仍进全篇级检查）
  added_ids[]/removed_ids[] —— 本轮新增段 / 上轮独有段（removed 只可出现在批注/说明中）
  protection.hits[]     —— 必须原样（禁改写）的段 + 命中规则；含 declared(上轮保护清单) 与
                            启发式规则（REF_LINE/CONFIG_LINE/ASCII_LINE/CAPTION/TOC/
                            LIST_ONLY/SUPSUB/FORMULA_HINT/SHORT）+ 元数据规则（OMML/RUN_SUP/
                            BOOKMARK/STYLE）+ --docx 规则（NUMPR 自动编号）
  global_watch.hits[]   —— 全篇级决策承载段（UNIT_STYLE/FIGTAB_REF/STANDARD_ID/LAW_ID/
                            BOOKTITLE）；**即使未变也必须全量重扫**
  suggested_shards[]    —— 建议分片：{name, count, chars, ids[]}，按文档序、按 --shard-chars 切
                            （增量模式应只 1~2 片）
  alignment             —— {mode, id_order_equal, channels{id,index,text,order,fuzzy},
                            id_hit_ratio, text_overlap_ratio, drift_segments[]}
                            id_hit_ratio 越低 id 漂移越严重；text_overlap_ratio<--min-overlap
                            说明可能不是相邻轮次
  naive                 —— 朴素口径（同 id / 同序号直连）的相同/不同段数，**只作漂移证据**
  token_hint            —— output_token_saving_ratio = 1 - 变化段字数占比
  warnings[]            —— id 漂移 / 段落增删 / 段数不等 / 重合率不足 等需人工确认的情况

退出码
  0 正常；2 参数或文件错误；3 有告警（id 漂移 / 段落增删 / 重合率 <--min-overlap），
  但结果已产出。**退出码 3 时不得直接把 changed_ids 当"本轮人工改动"用。**
""")
    ap.add_argument("--new", required=True, help="本轮新稿 paragraphs.json")
    ap.add_argument("--prev", help="上轮产物 merged_final.json / final.json / merged.json")
    ap.add_argument("--prev-dir", help="上轮工作目录；未给 --prev 时按 merged_final>final>merged 自动挑")
    ap.add_argument("--out", help="结果 JSON 输出路径")
    ap.add_argument("--md", help="markdown 摘要输出路径")
    ap.add_argument("--csv", help="把一行汇总追加到该 CSV")
    ap.add_argument("--protect", action="append", default=[], help="保护清单 json（可重复）")
    ap.add_argument("--meta", help="逐段元数据 json（fmt_scan.json 风格，含 omml/sups/bookmarks/style）")
    ap.add_argument("--docx", help="本轮源 docx（只读）：检测 numPr 自动编号 / OMML / run 级上下标")
    ap.add_argument("--align", choices=["auto", "text", "id"], default="auto",
                    help="auto=id 序一致则 id，否则文本锚定（默认）")
    ap.add_argument("--fuzzy", type=float, default=0.60, help="replace 块的相似度配对阈值（默认 0.60）")
    ap.add_argument("--ignore-space", action="store_true", help="比较前去掉所有空白再判定是否变化")
    ap.add_argument("--shard-chars", type=int, default=40000, help="建议分片的目标字数（默认 40000）")
    ap.add_argument("--min-overlap", type=float, default=0.95,
                    help="判定「确实相邻轮次」的文本重合率下限（默认 0.95）；低于则告警并返回 3")
    ap.add_argument("--json", action="store_true", help="把完整 JSON 打到 stdout")
    ap.add_argument("--quiet", action="store_true", help="不打印摘要")
    args = ap.parse_args(argv)

    if not args.prev and not args.prev_dir:
        ap.error("必须给 --prev 或 --prev-dir")
    if not os.path.exists(args.new):
        ap.error("--new 不存在: %s" % args.new)
    prev_pick = None
    if not args.prev:
        p, name = pick_prev(args.prev_dir)
        if not p:
            ap.error("--prev-dir 下找不到 %s 中任一个" % "/".join(PREV_PRIORITY))
        args.prev = p
        prev_pick = name
    if not os.path.exists(args.prev):
        ap.error("--prev 不存在: %s" % args.prev)

    res = build(args)
    if prev_pick:
        res["prev"]["picked_by"] = prev_pick

    if args.out:
        with io.open(args.out, "w", encoding="utf-8", newline="\n") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
    md = summary_md(res)
    if args.md:
        with io.open(args.md, "w", encoding="utf-8", newline="\n") as f:
            f.write("# round_diff 摘要\n\n" + md + "\n")
    if args.csv:
        new_row = not os.path.exists(args.csv)
        with io.open(args.csv, "a", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            if new_row:
                w.writerow(["new", "prev", "paras", "unchanged", "modified", "added", "removed",
                            "changed_ids", "changed_chars", "total_chars", "changed_char_ratio",
                            "id_order_equal", "protected_declared", "protected_hits",
                            "global_watch"])
            w.writerow([res["new"]["path"], res["prev"]["path"], res["counts"]["paras"],
                        res["counts"]["unchanged"], res["counts"]["modified"], res["counts"]["added"],
                        res["counts"]["removed"], res["counts"]["changed_ids"],
                        res["chars"]["changed"], res["chars"]["total"], res["chars"]["changed_ratio"],
                        res["alignment"]["id_order_equal"], res["protection"]["declared"],
                        len(res["protection"]["hits"]), len(res["global_watch"]["hits"])])
    if args.json:
        print(json.dumps(res, ensure_ascii=False, indent=1))
    if not args.quiet:
        print(md)
    if res["warnings"]:
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())
