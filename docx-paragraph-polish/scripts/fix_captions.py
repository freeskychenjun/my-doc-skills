# -*- coding: utf-8 -*-
"""fix_captions.py —— 题注编号自动修复（Word 修订模式写回副本）

读 check_figures_tables.py 的 figcheck.json，把可确定性修复的题注编号以修订标记改正：
  A级 PAIR_MISMATCH：正文引用编号为正确值，题注编号改成它（铁证）；
  B级 FMT_NO_DASH：丢连字符反推（尾段末位前插'-'，如 8.19→8.1-9），须满足
     反推后同 kind 内不重号、且同批反推构成从 -1 起的连续序列（或与 PAIR_MISMATCH 互证）；
  只建议不修：DUPLICATE_NUM（哪条该改属语义判断）、域生成的编号内空格（刷新会被覆盖）。

只改题注侧；错在题注、正文引用本就正确，题注改对后一致性自然恢复。
每处修复挂 Word 批注说明依据；输出 fix-report.json。

用法：
  python fix_captions.py <docx> <figcheck.json> [-o out.docx]
"""
import copy
import datetime
import json
import re
import shutil
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent))
import apply_track_changes as atc
from check_figures_tables import CAP_RE, norm_num, para_text

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def load_targets(figcheck, existing_nums):
    """从 figcheck.json + 现有题注全集 生成 {kind: {caption_num: (target_num, reason)}} 与建议清单。

    existing_nums: {kind: set(题注num)} —— 由本脚本先扫文档得到（figcheck 未存明细）。
    B 级序列校验：按前缀组合并「现有题注序号 + A级目标序号 + 本批反推序号」后必须从 1 连续；
    某前缀组不连续只降级该组，不影响其它组。
    """
    fix, suggest = [], []
    issues = figcheck.get("issues", [])
    dup = set()
    for i in issues:
        if i["category"] == "DUPLICATE_NUM":
            dup.add((i.get("kind"), i.get("num")))
        elif i["category"] == "SEQ_DUP":
            for n in i.get("dup", []):
                dup.add((i.get("kind"), "%s-%d" % (i["prefix"], n)))
    # A 级：PAIR_MISMATCH（正文引用编号为准）
    pair = {}
    for i in issues:
        if i["category"] == "PAIR_MISMATCH":
            pair.setdefault((i["kind"], i["caption_num"]), (
                i["num"], "正文引用为 %s%s，题注写作 %s%s" % (i["kind"], i["num"], i["kind"], i["caption_num"])))
    # B 级：FMT_NO_DASH 反推（"8.11" = 节"8.1"+序"1"，'-' 丢失）；已被 A 级覆盖的题注跳过
    guess = {}
    for i in issues:
        if i["category"] != "FMT_NO_DASH":
            continue
        num = i["num"]
        if (i.get("kind"), num) in pair:
            continue  # 同一题注已有 A 级铁证，不重复处理
        pre, tail = num.rsplit(".", 1)
        if len(tail) < 2 or not tail.isdigit():
            suggest.append((i.get("kind"), num, "无法反推（尾段不足 2 位）"))
            continue
        guess[(i["kind"], num)] = "%s.%s-%s" % (pre, tail[:-1], tail[-1])
    # 序列校验（按 kind×前缀 组独立）
    for kind in sorted({k for (k, _) in list(guess) + list(pair)}):
        pair_t = {t for (k2, _), (t, _) in pair.items() if k2 == kind}
        groups = {}
        for (k2, n), t in guess.items():
            if k2 != kind:
                continue
            groups.setdefault(t.rsplit("-", 1)[0], []).append((n, t))
        for prefix, items in groups.items():
            seqs = set()
            for t in pair_t:
                if t.startswith(prefix + "-") and t.rsplit("-", 1)[1].isdigit():
                    seqs.add(int(t.rsplit("-", 1)[1]))
            for num in existing_nums.get(kind, ()):
                if num.startswith(prefix + "-") and num.rsplit("-", 1)[1].isdigit():
                    seqs.add(int(num.rsplit("-", 1)[1]))
            conflict = False
            for n, t in items:
                s = t.rsplit("-", 1)[1]
                if (kind, t) in dup or t in pair_t or t in {x for x in existing_nums.get(kind, ())}:
                    suggest.append((kind, n, "反推为 %s 与现有/重复编号冲突" % t))
                    conflict = True
            if conflict:
                continue
            for _, t in items:
                seqs.add(int(t.rsplit("-", 1)[1]))
            if sorted(seqs) != list(range(1, max(seqs) + 1)):
                for n, t in items:
                    suggest.append((kind, n, "反推为 %s，但 %s- 序列不连续（现有 %s），需人工确认" % (
                        t, prefix, sorted(seqs))))
                continue
            for n, t in items:
                pair.setdefault((kind, n), (t, "丢连字符反推（%s→%s），与现有题注序列衔接自洽" % (n, t)))
    return pair, suggest


def iter_caption_paragraphs(doc):
    """产出 (p_element, order_or_None, match, full_text) —— 含表格内段落；para_text 已映射 noBreakHyphen"""
    body = doc.element.body
    order = 0
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            t = para_text(child)
            m = CAP_RE.match((t or "").strip()) if t and len(t.strip()) <= 90 else None
            if m and not (m.group(3) and re.search(r"[。；;]", m.group(3))):
                yield child, order, m, t
            order += 1
        elif child.tag == qn("w:tbl"):
            for p in child.iter(qn("w:p")):
                t = para_text(p)
                if not t or len(t.strip()) > 90:
                    continue
                m = CAP_RE.match(t.strip())
                if m and not (m.group(3) and re.search(r"[。；;]", m.group(3))):
                    yield p, None, m, t


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    args = sys.argv[1:]
    if len(args) < 2:
        print(__doc__)
        return 2
    doc_path = Path(args[0]).resolve()
    figcheck = json.loads(Path(args[1]).read_text(encoding="utf-8"))
    out_path = None
    if "-o" in args:
        out_path = Path(args[args.index("-o") + 1]).resolve()
    if out_path is None:
        out_path = doc_path.with_name(doc_path.stem + "-题注修订.docx")

    shutil.copyfile(doc_path, out_path)
    doc = Document(str(out_path))

    # 先扫全集（figcheck 未存题注明细）：{kind: set(num)}，用于 B 级序列衔接校验
    existing_nums = {}
    for p, order, m, full_text in iter_caption_paragraphs(doc):
        existing_nums.setdefault(m.group(1), set()).add(norm_num(m.group(2)))

    pair, suggest = load_targets(figcheck, existing_nums)

    atc.rev_date = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev = atc.next_rev_id(doc)

    report = {"fixed": [], "skipped": [], "suggest": [
        {"kind": k, "num": n, "reason": r} for k, n, r in suggest]}
    for p, order, m, full_text in iter_caption_paragraphs(doc):
        kind = m.group(1)
        num = norm_num(m.group(2))
        if (kind, num) not in pair:
            continue
        target, reason = pair[(kind, num)]
        has_field = ("fldChar" in p.xml) or ("instrText" in p.xml)
        if has_field:
            report["skipped"].append({"kind": kind, "num": num,
                "reason": "题注含 Word 域（自动编号），编号格式由域代码决定（如 STYLEREF+SEQ 无连字符）；改文本会被域刷新覆盖——需改域代码或先转固定文本，交人工"})
            continue
        # 构造新文本：仅替换编号子串（用整段重新匹配拿 span）
        core = full_text.strip()
        cm = CAP_RE.match(core)
        if not cm:
            report["skipped"].append({"kind": kind, "num": num, "reason": "重新匹配失败"})
            continue
        new_core = core[:cm.start(2)] + target + core[cm.end(2):]
        if new_core == core:
            continue
        chars = atc.collect_chars(p)
        if "".join(c for c, _ in chars).strip() != core:
            report["skipped"].append({"kind": kind, "num": num, "reason": "段落含非 w:t 内容，跳过"})
            continue
        prefix = full_text[:len(full_text) - len(full_text.lstrip())]
        suffix = full_text[len(full_text.rstrip()):]
        core_chars = chars[len(prefix):len(chars) - len(suffix)] if suffix else chars[len(prefix):]
        ops = atc.char_diff(core, new_core)
        if ops is None:
            report["skipped"].append({"kind": kind, "num": num, "reason": "diff 失败"})
            continue
        try:
            rev = atc.rewrite_paragraph(p, ops, core_chars, rev)
            if prefix or suffix:
                atc.restore_ws(p, prefix, suffix, core_chars)
            atc.add_flag_comment(doc, p, "题注编号自动修正：%s%s → %s%s（%s）。依据：%s" % (
                kind, num, kind, target, reason, figcheck.get("file", "")))
            report["fixed"].append({"kind": kind, "old": num, "new": target,
                                    "reason": reason, "order": order})
        except Exception as e:
            report["skipped"].append({"kind": kind, "num": num, "reason": "写入失败: %r" % e})

    doc.save(str(out_path))
    rp = out_path.with_name(out_path.stem + ".fix-report.json")
    rp.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print("fixed=%d skipped=%d suggest=%d" % (len(report["fixed"]), len(report["skipped"]), len(report["suggest"])))
    for f in report["fixed"]:
        print("  FIX %s%s -> %s%s (%s)" % (f["kind"], f["old"], f["kind"], f["new"], f["reason"][:40]))
    for s in report["skipped"]:
        print("  SKIP %s%s: %s" % (s["kind"], s["num"], s["reason"]))
    for s in report["suggest"]:
        print("  SUGGEST %s%s: %s" % (s["kind"], s["num"], s["reason"]))
    print("out=%s" % out_path)
    print("report=%s" % rp)
    return 0


if __name__ == "__main__":
    sys.exit(main())
