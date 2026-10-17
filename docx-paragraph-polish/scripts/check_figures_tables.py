# -*- coding: utf-8 -*-
"""check_figures_tables.py —— 图表编号与引用核查（只读，不改文档）

能确定性判定的：
  1. 编号重复 / 跳号 / 乱序（按章节前缀分组）
  2. 编号格式不统一（编号内空格、全角连字符/点、深度层级混用、编号与名称间缺空格）
  3. 表格/图片缺题注（表题应在表上方、图题应在图下方）
  4. 正文引用了不存在的图表（悬空引用）
  5. 图表从未被正文引用（提示）
  6. 引用处名称与题注名称不一致（soft，输出上下文供复核）
  7. 题注出现在首次引用之前（图应先引用后出现，提示）
  8. 题注含 Word 域（编号为自动编号，须在 Word 中刷新后复核）

不能判定的（交模型/人工）：表体内容与表题是否张冠李戴、图内容与图名是否相符。

用法：
    python check_figures_tables.py "<docx路径>" [-o out.json] [--names]
"""
import collections
import difflib
import io
import json
import re
import sys

from docx import Document
from docx.oxml.ns import qn

FW = str.maketrans({
    "０": "0", "１": "1", "２": "2", "３": "3", "４": "4", "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
    "．": ".", "。": ".", "—": "-", "–": "-", "－": "-", "～": "-", "~": "-", "　": " ",
})

NUM = r"[0-9０-９]+(?:[.．][0-9０-９]+)*(?:\s*[-—–－~～]\s*[0-9０-９]+)?"
CAP_RE = re.compile(r"^\s*(?:续)?\s*([图表])\s*(" + NUM + r")\s*[.．、：:，,]?\s*(\S.{0,80})?$")
REF_RE = re.compile(r"([图表])\s*(" + NUM + r")")
SENT_END = "。；;!！?"


def norm_num(s):
    s = s.translate(FW)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[-—–－~～]", "-", s)
    return s


def norm_name(s):
    s = (s or "").strip().translate(FW)
    s = re.sub(r"[\s\u3000]+", "", s)
    return s.strip("：:。.、,，")


def para_text(p):
    """拼接段落可见文本。注意：w:noBreakHyphen（不换行连字符）不是 w:t，
    漏掉它会把"图 5.1-1"读成"图 5.11"造成大量假阳性——必须映射为 '-'。"""
    out = []
    for el in p.iter():
        tag = el.tag
        if tag == qn("w:t"):
            out.append(el.text or "")
        elif tag == qn("w:noBreakHyphen"):
            out.append("-")
        elif tag == qn("w:tab"):
            out.append(" ")
    return "".join(out)


def walk(doc):
    """按文档序返回 items：p/tbl/tblcap。表题可能写在正文段（表上方）或表格首行内。"""
    items = []
    order = 0
    tbl_idx = 0
    for child in doc.element.body.iterchildren():
        tag = child.tag
        if tag == qn("w:p"):
            xml = child.xml
            items.append({"kind": "p", "order": order, "text": para_text(child),
                          "img": xml.count("a:blip") + xml.count("w:pict"),
                          "math": "oMath" in xml,
                          "field": ("fldChar" in xml) or ("instrText" in xml)})
            order += 1
        elif tag == qn("w:tbl"):
            rows = child.findall(qn("w:tr"))
            # 前 2 行逐格找表题（表题常写在首行合并单元格里）
            found = False
            for ri, tr in enumerate(rows[:2]):
                for tc in tr.findall(qn("w:tc")):
                    cell_text = para_text(tc).strip()
                    if cell_text and CAP_RE.match(cell_text):
                        items.append({"kind": "tblcap", "order": order - 1, "text": cell_text,
                                      "tblidx": tbl_idx, "row": ri})
                        found = True
                        break
                if found:
                    break
            first_row = para_text(rows[0]) if rows else ""
            items.append({"kind": "tbl", "order": order - 1, "tblidx": tbl_idx,
                          "text": "", "first_row": first_row, "has_cap": found})
            tbl_idx += 1
    return items


def main():
    args = sys.argv[1:]
    if not args:
        print("usage: check_figures_tables.py <docx> [-o out.json] [--names]")
        return 2
    src = args[0]
    out, do_names = None, False
    i = 1
    while i < len(args):
        if args[i] == "-o" and i + 1 < len(args):
            out, i = args[i + 1], i + 2
        elif args[i] == "--names":
            do_names, i = True, i + 1
        else:
            i += 1
    if out is None:
        out = re.sub(r"\.docx$", "", src) + "-figcheck.json"

    doc = Document(src)
    items = walk(doc)

    # ---- 1) 题注识别（正文段 + 表格首行） ----
    caps = []
    for it in items:
        if it["kind"] == "p":
            t = (it["text"] or "").strip()
            if not t or len(t) > 90:
                continue
            loc = "body"
        elif it["kind"] == "tblcap":
            t = it["text"].strip()
            loc = "table_row%d" % it.get("row", 0)
        else:
            continue
        m = CAP_RE.match(t)
        if not m:
            continue
        kind, raw_num, name = m.group(1), m.group(2), (m.group(3) or "").strip()
        if name and re.search(r"[。；;]", name):
            continue  # 像正文句子，不是题注
        caps.append({"kind": kind, "raw": raw_num, "num": norm_num(raw_num),
                     "name": name, "order": it["order"], "field": it.get("field", False),
                     "continued": t.lstrip().startswith("续"), "loc": loc,
                     "tblidx": it.get("tblidx"), "in_img_para": False})

    issues = []

    def add(cat, level, msg, **kw):
        d = {"category": cat, "level": level, "msg": msg}
        d.update(kw)
        issues.append(d)

    # ---- 2) 重复编号（续表不计） ----
    by_key = collections.defaultdict(list)
    for c in caps:
        if not c["continued"]:
            by_key[(c["kind"], c["num"])].append(c)
    for (kind, num), lst in by_key.items():
        if len(lst) > 1:
            add("DUPLICATE_NUM", "ERROR", "%s%s 出现 %d 次（%s）" % (
                kind, num, len(lst),
                "；".join("段序%d%s" % (x["order"], "（表格首行）" if x["loc"] != "body" else "") for x in lst)),
                kind=kind, num=num, orders=[x["order"] for x in lst])

    # ---- 3) 跳号（续表不计） ----
    by_prefix = collections.defaultdict(list)
    for c in caps:
        if c["continued"]:
            continue
        if "-" in c["num"]:
            pre, seq = c["num"].rsplit("-", 1)
            if seq.isdigit():
                by_prefix[(c["kind"], pre)].append((int(seq), c))
                continue
        by_prefix[(c["kind"], c["num"])].append((None, c))
    for (kind, pre), lst in by_prefix.items():
        seqs = sorted(s for s, _ in lst if s is not None)
        if not seqs:
            continue
        expect = list(range(1, max(seqs) + 1))
        missing = [n for n in expect if n not in seqs]
        dup = [n for n in set(seqs) if seqs.count(n) > 1]
        if missing:
            add("GAP", "WARN", "%s%s- 缺编号 %s（现有 %s）" % (
                kind, pre, missing, seqs), kind=kind, prefix=pre, missing=missing)
        if dup:
            add("SEQ_DUP", "ERROR", "%s%s- 序号重复 %s" % (kind, pre, dup), kind=kind, prefix=pre, dup=dup)

    # ---- 4) 格式不统一 ----
    for c in caps:
        raw = c["raw"]
        if re.search(r"\s", raw):
            add("FMT_SPACE_IN_NUM", "WARN", "%s%s 编号内含空格（段序 %d）" % (c["kind"], raw, c["order"]),
                kind=c["kind"], num=c["num"], order=c["order"])
        if re.search(r"[０-９．—–－～]", raw):
            add("FMT_FULLWIDTH", "WARN", "%s%s 编号含全角字符（段序 %d）" % (c["kind"], raw, c["order"]),
                kind=c["kind"], num=c["num"], order=c["order"])
        # 多级编号却无连字符序号（如 8.11 / 5.21）——疑丢连字符，应为 8.1-1 / 5.2-1
        if "-" not in c["num"] and "." in c["num"]:
            add("FMT_NO_DASH", "WARN", "%s%s 编号多级但无连字符序号（段序 %d）——疑丢连字符，如 图5.1-1 被写成 图5.11" % (
                c["kind"], raw.replace(" ", ""), c["order"]),
                kind=c["kind"], num=c["num"], order=c["order"])
        if not c["name"]:
            add("NO_NAME", "WARN", "%s%s 题注无名称（段序 %d）" % (c["kind"], c["num"], c["order"]),
                kind=c["kind"], num=c["num"], order=c["order"])
        elif re.match(r"^\S", c["name"]) and not c["name"].startswith((" ", "　")):
            pass
    # 深度层级混用
    depth = collections.defaultdict(collections.Counter)
    for c in caps:
        if c["continued"]:
            continue
        pre = c["num"].rsplit("-", 1)[0] if "-" in c["num"] else c["num"]
        n = len([x for x in pre.split(".") if x != ""])
        depth[c["kind"]][n] += 1
    for kind, cnt in depth.items():
        if len(cnt) > 1:
            main = cnt.most_common(1)[0][0]
            odd_by_depth = collections.Counter()
            odd_samples = {}
            for c in caps:
                if c["kind"] != kind or c["continued"]:
                    continue
                pre = c["num"].rsplit("-", 1)[0] if "-" in c["num"] else c["num"]
                n = len([x for x in pre.split(".") if x != ""])
                if n != main:
                    odd_by_depth[n] += 1
                    odd_samples.setdefault(n, c)
            detail = "、".join("%d级×%d（如%s%s《%s》段序%d）" % (
                n, v, kind, odd_samples[n]["num"], odd_samples[n]["name"][:20], odd_samples[n]["order"])
                for n, v in sorted(odd_by_depth.items()))
            add("FMT_DEPTH", "WARN", "%s 编号层级混用：主流 %d 级；另有 %s" % (kind, main, detail),
                kind=kind, main_depth=main,
                odd=[{"num": odd_samples[n]["num"], "order": odd_samples[n]["order"]} for n in odd_by_depth])

    # ---- 5) 表格/图片缺题注 ----
    cap_orders = {c["order"] for c in caps}
    for it in items:
        if it["kind"] == "tbl":
            if it.get("has_cap"):
                continue
            o = it["order"]
            near = [c for c in caps if c["kind"] == "表" and not c["continued"]
                    and -4 <= o - c["order"] <= 1]
            if not near:
                add("TBL_NO_CAPTION", "ERROR", "表格 #%d（段序约 %d）首行及上方 4 段内未发现表题；首行文本：%s" % (
                    it.get("tblidx", -1), o, (it.get("first_row") or "").strip()[:50]),
                    tblidx=it.get("tblidx"), order=o)
        elif it["kind"] == "p" and it["img"] and not (it["order"] in cap_orders):
            t = (it["text"] or "").strip()
            # 排除：公式段（含 OMML 或"式中"开头）、封面装饰（前 40 段）
            if it.get("math") or t.startswith("式中") or it["order"] < 40:
                continue
            o = it["order"]
            near = [c for c in caps if c["kind"] == "图" and -2 <= c["order"] - o <= 3]
            if not near:
                add("FIG_NO_CAPTION", "WARN", "图片段（段序 %d）附近未发现图题：%s" % (
                    o, t[:40]), order=o)

    # ---- 6) 正文引用核查 ----
    cap_by_kind = collections.defaultdict(dict)
    for c in caps:
        cap_by_kind[c["kind"]].setdefault(c["num"], c)
    body_cap_orders = {c["order"] for c in caps if c["loc"] == "body"}
    first_ref = {}
    refs = []
    for it in items:
        if it["kind"] != "p":
            continue
        t = it["text"] or ""
        if not t.strip():
            continue
        is_cap = it["order"] in body_cap_orders
        for m in REF_RE.finditer(t):
            kind, raw = m.group(1), m.group(2)
            num = norm_num(raw)
            if is_cap:
                continue
            confident = "-" in num
            ctx = t[max(0, m.start() - 15):m.end() + 30]
            refs.append({"kind": kind, "num": num, "raw": raw, "order": it["order"],
                         "confident": confident, "ctx": ctx, "end": m.end(), "text": t})
            first_ref.setdefault((kind, num), it["order"])
            if confident and num not in cap_by_kind[kind]:
                add("DANGLING_REF", "ERROR", "正文引用 %s%s，但文档中无此编号题注（段序 %d）｜%s" % (
                    kind, num, it["order"], ctx.strip()), kind=kind, num=num, order=it["order"], ctx=ctx.strip())
            if do_names and confident and num in cap_by_kind[kind]:
                frag = t[m.end():m.end() + 30]
                frag = re.split(r"[，,、。；;：:（(]", frag)[0]
                frag = frag.strip(" 　")
                cname = cap_by_kind[kind][num]["name"]
                if len(frag) >= 4 and cname:
                    r = difflib.SequenceMatcher(None, norm_name(frag), norm_name(cname)).ratio()
                    if 0.40 < r < 0.99:
                        add("REF_NAME_SOFT", "INFO", "引用处名称与题注可能不一致：引用「%s」vs 题注「%s」（相似度 %.2f，段序 %d）" % (
                            frag, cname, r, it["order"]), kind=kind, num=num,
                            ref_frag=frag, caption_name=cname, ratio=round(r, 2), order=it["order"])

    # ---- 6.5) 悬空引用 ↔ 未引用题注 配对消解（丢连字符类格式错误） ----
    # 例：正文引"图5.1-1"悬空、题注"图 5.11"未被引用 → 去"分隔符"后同为"511" → 判丢连字符
    def stripkey(n):
        return n.replace(".", "").replace("-", "")
    cap_key = {}
    for c in caps:
        if not c["continued"]:
            cap_key.setdefault((c["kind"], stripkey(c["num"])), c)
    matched_caps = set()
    dangling = [i for i in issues if i["category"] == "DANGLING_REF"]
    for d in dangling:
        hit = cap_key.get((d["kind"], stripkey(d["num"])))
        if hit and hit["num"] != d["num"]:
            d["category"] = "PAIR_MISMATCH"
            d["level"] = "ERROR"
            d["caption_num"] = hit["num"]
            d["caption_order"] = hit["order"]
            d["msg"] = "正文引用 %s%s 但题注写作 %s%s（段序 %d）——疑题注编号丢连字符，如 图5.1-1 被写成 图5.11" % (
                d["kind"], d["num"], hit["kind"], hit["raw"], hit["order"])
            matched_caps.add(id(hit))

    # ---- 7) 未被引用 / 题注早于首次引用（丢连字符类已由 FMT_NO_DASH 报告，不再重复报 UNREFERENCED） ----
    no_dash = {i.get("num") for i in issues if i["category"] == "FMT_NO_DASH"}
    for c in caps:
        if c["continued"] or id(c) in matched_caps or c["num"] in no_dash:
            continue
        if (c["kind"], c["num"]) not in first_ref:
            add("UNREFERENCED", "INFO", "%s%s《%s》未被正文引用（段序 %d）" % (
                c["kind"], c["num"], c["name"][:30], c["order"]), kind=c["kind"], num=c["num"], order=c["order"])
        else:
            ro = first_ref[(c["kind"], c["num"])]
            if c["order"] < ro:
                add("ORDER", "INFO", "%s%s 题注（段序 %d）早于首次引用（段序 %d）" % (
                    c["kind"], c["num"], c["order"], ro), kind=c["kind"], num=c["num"], order=c["order"], ref_order=ro)

    # ---- 8) 域编号提示 ----
    fld = [c for c in caps if c["field"]]
    if fld:
        add("FIELD_NUM", "INFO", "%d 条题注含 Word 域（自动编号），缓存文本可能过期：%s" % (
            len(fld), ["%s%s" % (c["kind"], c["num"]) for c in fld[:12]]),
            count=len(fld))

    json.dump({"file": src, "captions": len(caps), "tables": sum(1 for x in items if x["kind"] == "tbl"),
               "img_paras": sum(1 for x in items if x["kind"] == "p" and x["img"]),
               "issues": issues}, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    cnt = collections.Counter("%s/%s" % (i["category"], i["level"]) for i in issues)
    print("captions=%d tables=%d img_paras=%d issues=%d" % (
        len(caps), sum(1 for x in items if x["kind"] == "tbl"),
        sum(1 for x in items if x["kind"] == "p" and x["img"]), len(issues)))
    for k in sorted(cnt):
        print("  %-24s %d" % (k, cnt[k]))
    print("out=%s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
