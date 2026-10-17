# -*- coding: utf-8 -*-
"""check_refs.py —— 法规/标准引用确定性核查（docx-paragraph-polish 第 6 步查证遍的 A/B 类）

把段落里对法律法规、技术标准的引用，与两份本机权威对照表比对，直接给出裁决建议：
  - 名称与编号张冠李戴（《堤防工程设计规范》（GB50286-2013）实际是《防洪标准》）
  - 引用旧版本年份（表内已有现行版本）
  - 法规令号/年份与名称不符（《河道管理条例》（2018年国务院令第829号））

用法：
    python check_refs.py "<paragraphs.json>" [-o out.json] [--laws <法律法规.md>] [--stds <水利标准.md>]

输出 <paragraphs>-refcheck.json；stdout 只打印统计与路径（避免 GBK 乱码）。
只做"可确定性判定"的部分：表中查不到的引用一律标 NOT_FOUND（不算问题，交模型/人工）。
"""
import json
import os
import re
import sys

DEFAULT_LAWS = r"D:\Claude Code\报告智能检查skill\references\法律法规.md"
DEFAULT_STDS = r"D:\Claude Code\报告智能检查skill\references\水利标准.md"

FW = str.maketrans({
    "０": "0", "１": "1", "２": "2", "３": "3", "４": "4", "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
    "Ａ": "A", "Ｂ": "B", "Ｃ": "C", "Ｄ": "D", "Ｅ": "E", "Ｆ": "F", "Ｇ": "G", "Ｈ": "H", "Ｉ": "I", "Ｊ": "J",
    "Ｋ": "K", "Ｌ": "L", "Ｍ": "M", "Ｎ": "N", "Ｏ": "O", "Ｐ": "P", "Ｑ": "Q", "Ｒ": "R", "Ｓ": "S", "Ｔ": "T",
    "Ｕ": "U", "Ｖ": "V", "Ｗ": "W", "Ｘ": "X", "Ｙ": "Y", "Ｚ": "Z", "（": "(", "）": ")", "／": "/",
    "—": "-", "–": "-", "－": "-", "~": "-", "～": "-", "．": ".", "。": ".", "　": " ",
})

STD_SERIES = ("GB/T", "GB", "SL/T", "SL", "NB/T", "NB", "DL/T", "DL", "CJJ", "CJJ/T", "JGJ", "JTG", "JTS",
              "CH/T", "CH", "DB36", "DB/T", "JT/T", "QX/T", "HJ", "YD/T", "CECS")


def norm(s):
    return re.sub(r"\s+", "", (s or "").translate(FW).replace("'", "").replace("’", ""))


def split_code(code):
    """'SL483-2017' -> ('SL483', '2017')；兼容 SL/T278-2020、GB 73.7-2013"""
    m = re.match(r"^([A-Za-z/]*\d+(?:\.\d+)*)-(\d{4})$", code)
    if m:
        return m.group(1), m.group(2)
    return code, ""


def base_series(base):
    """SL/T278 -> ('SL278','T')；GB/T14511 -> ('GB14511','T')"""
    m = re.match(r"^([A-Za-z]+)/T(\d+(?:\.\d+)*)$", base)
    if m:
        return m.group(1) + m.group(2), "T"
    m = re.match(r"^([A-Za-z]+)(\d+(?:\.\d+)*)$", base)
    if m:
        return m.group(1) + m.group(2), ""
    return base, ""


def read_table(path):
    rows = []
    for line in open(path, encoding="utf-8"):
        line = line.rstrip("\n")
        if not re.match(r"^\|\s*\d+\s*\|", line):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) >= 4:
            rows.append(cells[1:])
    return rows


class Refs(object):
    def name_in_tables(self, name):
        """书名号里的名称自身是否就是对照表中的某个标准/法规名"""
        n = norm(name)
        for k in (n, re.sub(r"\((?:修订|修正|试行|暂行)\)", "", n)):
            if k in self.std_by_name or k in self.law_by_name:
                return True
        return False
    def __init__(self, laws_path, stds_path):
        # 标准：编号 -> 记录；基准号 -> 记录列表
        self.std_by_code = {}
        self.std_by_base = {}
        self.std_by_name = {}
        for code_raw, name, date in [(r[0], r[1], r[2]) for r in read_table(stds_path)]:
            code = norm(code_raw)
            if not code or code.lower() == "nan":
                continue
            base, year = split_code(code)
            rec = {"code": code, "codeRaw": code_raw, "name": name, "date": date, "year": year}
            self.std_by_code.setdefault(code, rec)
            b, t = base_series(base)
            self.std_by_base.setdefault(b, []).append(rec)
            self.std_by_name.setdefault(norm(name), []).append(rec)
        # 法规：名称 -> 记录；令号 -> 记录
        self.law_by_name = {}
        self.law_by_docnum = {}
        for docnum, name, date, note in [(r[0], r[1], r[2], r[3] if len(r) > 3 else "") for r in read_table(laws_path)]:
            rec = {"docnum": docnum, "name": name, "date": date, "note": note}
            for key in self._law_keys(name):
                self.law_by_name.setdefault(key, []).append(rec)
            dn = self._docnum_key(docnum)
            if dn:
                self.law_by_docnum.setdefault(dn, []).append(rec)

    @staticmethod
    def _law_keys(name):
        n = norm(name)
        n = re.sub(r"\((?:修订|修正|制定|20\d\d年修订|20\d\d年修正|试行|暂行)?\)", "", n)
        n = re.sub(r"(20\d\d年修订|20\d\d年修正)$", "", n)
        keys = {n}
        if n.startswith("中华人民共和国"):
            keys.add(n[len("中华人民共和国"):])
        keys.discard("")
        return keys

    @staticmethod
    def _docnum_key(docnum):
        s = norm(docnum)
        if not s or s.lower() == "nan" or "令" not in s:
            return ""
        year = re.search(r"(19|20)\d\d", s)
        num = re.search(r"第?(\d+)号", s)
        if not num:
            return ""
        issuer = "主席令" if "主席令" in s else ("国务院令" if "国务院令" in s else "国务院")
        return "%s|%s|%s" % (issuer, year.group(0) if year else "", num.group(1))


BOOK = re.compile(r"《([^》]{2,60})》")
CITE_CODE = re.compile(r"^\s*[（(]?\s*((?:%s)(?:/T)?\s*\d+(?:\.\d+)*\s*[-—–~－]?\s*\d{0,4})\s*[）)]?" %
                       "|".join(s.replace("/", r"/") for s in STD_SERIES))
ANY_CODE = re.compile(r"(?:%s)(?:/T)?\s*\d+(?:\.\d+)*(?:\s*[-—–~－]\s*\d{4})?" %
                      "|".join(s.replace("/", r"/") for s in STD_SERIES))
DOCNUM_IN_TEXT = re.compile(r"((?:19|20)\d\d)\s*年?\s*(主席令|国务院令|国务院)\s*(?:第)?\s*(\d+)\s*号")


def name_matches(a, b):
    """书名号里的名称与表中名称是否同指（容忍简称与括号后缀差异）"""
    na, nb = norm(a), norm(b)
    na = re.sub(r"\((?:修订|修正|试行|暂行|20\d\d年?(?:修订|修正)?)\)", "", na)
    nb = re.sub(r"\((?:修订|修正|试行|暂行|20\d\d年?(?:修订|修正)?)\)", "", nb)
    if na == nb:
        return True
    for pre in ("中华人民共和国",):
        if na.startswith(pre):
            na = na[len(pre):]
        if nb.startswith(pre):
            nb = nb[len(pre):]
    return na == nb


def check_paragraph(text, refs, pid, para_index):
    hits = []

    def mismatch(name, rec):
        """名称与编号不符：区分「名称是表中标准名 → 改编号」与「名称非标准名 → 引用本身可疑」"""
        if refs.name_in_tables(name):
            alt = (refs.std_by_name.get(norm(name)) or [{}])[0]
            return {"status": "NAME_CODE_MISMATCH", "action": "FIX_CODE",
                    "expected": alt.get("codeRaw", ""),
                    "note": "名称与编号不符：编号 %s 在对照表中为《%s》；文中《%s》本身是表中标准名，应按名称把编号改为 %s" % (
                        rec["codeRaw"], rec["name"], name, alt.get("codeRaw", "?")), "table": rec}
        return {"status": "NAME_CODE_MISMATCH", "action": "VERIFY_CITATION", "expected": "",
                "note": "名称与编号不符：编号 %s 在对照表中为《%s》，而《%s》不在对照表任何编号下（疑非标准名或误记）"
                        "——勿机械按编号改名称，须按正文主题核实该引用本身" % (rec["codeRaw"], rec["name"], name),
                "table": rec}

    # --- 1) 《名称》（编号或令号） ---
    for m in BOOK.finditer(text):
        name = m.group(1).strip()
        tail = text[m.end():m.end() + 40]
        code_m = CITE_CODE.match(tail)
        code = norm(code_m.group(1)) if code_m else ""
        # 1a) 名称 × 标准编号
        if code:
            base, year = split_code(code)
            b, series = base_series(base)
            rec = refs.std_by_code.get(code)
            status, note, expected, action = None, "", "", ""
            if rec and name_matches(name, rec["name"]):
                status = "OK"
            elif rec and not name_matches(name, rec["name"]):
                mm = mismatch(name, rec)
                status, note, expected, action = mm["status"], mm["note"], mm["expected"], mm["action"]
            else:
                cands = refs.std_by_base.get(b, [])
                if cands:
                    by_year = [c for c in cands if c["year"] == year]
                    same_name = [c for c in cands if name_matches(name, c["name"])]
                    if by_year and not name_matches(name, by_year[0]["name"]):
                        mm = mismatch(name, by_year[0])
                        status, note, expected, action = mm["status"], mm["note"], mm["expected"], mm["action"]
                    elif same_name and not by_year:
                        cur = same_name[0]
                        status, note = "OLD_VERSION", "文中年份 %s 与对照表现行版本 %s（%s）不一致" % (
                            year or "缺", cur["codeRaw"], cur["date"])
                        expected, action = cur["codeRaw"], "FIX_CODE"
                    else:
                        status = "NOT_FOUND"
                else:
                    status = "NOT_FOUND"
            if status != "OK":
                hits.append({"id": pid, "paraIndex": para_index, "kind": "std",
                             "citeName": name, "citeCode": code_m.group(1).strip() if code_m else "",
                             "normCode": code, "status": status, "action": action, "expected": expected,
                             "note": note, "table": (rec or (refs.std_by_base.get(b) or [{}])[0])})
        # 1b) 名称 × 法规令号
        dn = DOCNUM_IN_TEXT.search(tail) or DOCNUM_IN_TEXT.search(text[m.end():m.end() + 80])
        if dn and not code:
            year, issuer, num = dn.group(1), dn.group(2), dn.group(3)
            key = "%s|%s|%s" % (issuer, year, num)
            key_noyear = "%s|%s" % (issuer, num)
            recs = refs.law_by_name.get(norm(name)) or []
            if not recs:
                for k in Refs._law_keys(name):
                    recs = refs.law_by_name.get(k) or []
                    if recs:
                        break
            if recs:
                ok = any(Refs._docnum_key(r["docnum"]) == key for r in recs)
                wrong_year = any(Refs._docnum_key(r["docnum"]).startswith(key_noyear) for r in recs)
                if not ok:
                    cur = recs[0]
                    status = "LAW_DOCNUM_MISMATCH"
                    note = "文中令号「%s年%s第%s号」与对照表「%s」不符" % (year, issuer, num, cur["docnum"])
                    hits.append({"id": pid, "paraIndex": para_index, "kind": "law",
                                 "citeName": name, "citeCode": "%s年%s第%s号" % (year, issuer, num),
                                 "status": status, "action": "FIX_DOCNUM", "expected": cur["docnum"],
                                 "note": note + ("（年份有误）" if wrong_year else ""), "table": cur})
    # --- 2) 段落中孤立标准号：基准号在表中但年份非现行 ---
    for m in ANY_CODE.finditer(text):
        raw = m.group(0)
        if m.start() > 0 and text[m.start() - 1] in "《》":
            continue
        code = norm(raw)
        if not re.search(r"\d{4}$", code):
            continue
        base, year = split_code(code)
        b, _ = base_series(base)
        cands = refs.std_by_base.get(b, [])
        if not cands:
            continue
        if refs.std_by_code.get(code):
            continue  # 完全命中，无需报
        same_year = [c for c in cands if c["year"] == year]
        if same_year:
            continue
        cur = cands[0]
        ctx_book = None
        for bm in BOOK.finditer(text):
            if bm.end() <= m.start() and m.start() - bm.end() < 40:
                ctx_book = bm.group(1)
        if ctx_book and not name_matches(ctx_book, cur["name"]):
            continue  # 名称—编号组合已在 1) 里判过
        hits.append({"id": pid, "paraIndex": para_index, "kind": "std",
                     "citeName": ctx_book or "", "citeCode": raw.strip(), "normCode": code,
                     "status": "STD_NOT_CURRENT" if year else "NOT_FOUND",
                     "action": "FIX_CODE", "expected": cur["codeRaw"],
                     "note": "基准号 %s 在对照表中为 %s（%s），文中写作 %s" % (b, cur["codeRaw"], cur["date"], raw.strip()),
                     "table": cur})
    return hits


def main():
    args = sys.argv[1:]
    if not args:
        print("usage: check_refs.py <paragraphs.json> [-o out.json] [--laws f] [--stds f]")
        return 2
    src = args[0]
    out = None
    laws, stds = DEFAULT_LAWS, DEFAULT_STDS
    include_all = False
    i = 1
    while i < len(args):
        if args[i] == "-o" and i + 1 < len(args):
            out = args[i + 1]
            i += 2
        elif args[i] == "--laws" and i + 1 < len(args):
            laws = args[i + 1]
            i += 2
        elif args[i] == "--stds" and i + 1 < len(args):
            stds = args[i + 1]
            i += 2
        elif args[i] == "--all":
            include_all = True
            i += 1
        else:
            i += 1
    if out is None:
        out = re.sub(r"\.json$", "", src) + "-refcheck.json"

    refs = Refs(laws, stds)
    data = json.load(open(src, encoding="utf-8"))
    all_hits = []
    for it in data:
        text = it.get("text", "") or ""
        if "《" not in text and not ANY_CODE.search(text):
            continue
        all_hits += check_paragraph(text, refs, it.get("id"), it.get("paraIndex"))

    # 去重：同一段同一编号只留一条（按 名称不符 > 旧版本 > 非现行 的优先序）
    prio = {"NAME_CODE_MISMATCH": 0, "LAW_DOCNUM_MISMATCH": 0, "OLD_VERSION": 1,
            "STD_NOT_CURRENT": 2, "NOT_FOUND": 3}
    best = {}
    for h in all_hits:
        k = (h["id"], h.get("normCode") or h.get("citeCode"))
        if k not in best or prio.get(h["status"], 9) < prio.get(best[k]["status"], 9):
            best[k] = h
    all_hits = list(best.values())
    if not include_all:
        all_hits = [h for h in all_hits if h["status"] != "NOT_FOUND"]
    all_hits.sort(key=lambda h: (h.get("paraIndex") or 0, h["status"]))

    json.dump(all_hits, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    counts = {}
    for h in all_hits:
        counts[h["status"]] = counts.get(h["status"], 0) + 1
    print("paragraphs=%d  hits=%d%s" % (len(data), len(all_hits), "" if include_all else "  (NOT_FOUND 已抑制，--all 可显示)"))
    for k in sorted(counts):
        print("  %-22s %d" % (k, counts[k]))
    print("out=%s" % out)
    print("table_rows: laws=%d stds=%d" % (
        len([r for r in read_table(laws)]), len([r for r in read_table(stds)])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
