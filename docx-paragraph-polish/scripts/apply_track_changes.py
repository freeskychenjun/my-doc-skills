# -*- coding: utf-8 -*-
"""以 Word 修订(track changes)模式把逐段改写写回 docx 副本。

纯 OpenXML 实现（python-docx + lxml），不依赖本机安装 Word/WPS；
产物是带 w:ins/w:del 修订标记的标准 docx，Word 与 WPS 均可逐条接受/拒绝。

移植自 WordTools/SmartWriting/{CharDiff,TrackChangesRewriter}.cs：
- 字符级 LCS diff，删字标删除修订、增字标插入修订，未改动部分原样保留；
- 净改动占比超门槛（润色 0.6 / 逻辑重构 0.95，硬顶 0.95）自动跳过；
- 段首 sanity 校验：段落文本与 extract 时不一致则跳过。

用法:
  python apply_track_changes.py <docx> <rewrites.json> [polish|logic] [输出docx路径]
rewrites.json: [{"id":1,"paraIndex":3,"original":"...","newText":"..."}, ...]
（paraIndex/original 必须来自 extract_body_paragraphs.py 的输出，不得手改）
"""
import copy
import json
import locale
import re
import shutil
import sys
import datetime
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
AUTHOR = 'AI校正'
HARD_CEILING = 0.95
MODE_CEILING = {'polish': 0.6, 'logic': 0.95}
MAX_MATRIX = 4_000_000


def char_diff(a: str, b: str):
    """标准 LCS diff，返回 [(op, text), ...]，op in {equal,delete,insert}；输入过大返回 None。"""
    if not a and not b:
        return []
    if not a:
        return [('insert', b)]
    if not b:
        return [('delete', a)]
    n, m = len(a), len(b)
    if n * m > MAX_MATRIX:
        return None
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        ai = a[i - 1]
        row, prev = dp[i], dp[i - 1]
        for j in range(1, m + 1):
            if ai == b[j - 1]:
                row[j] = prev[j - 1] + 1
            else:
                row[j] = row[j - 1] if row[j - 1] >= prev[j] else prev[j]
    ops = []
    ii, jj = n, m
    while ii > 0 and jj > 0:
        if a[ii - 1] == b[jj - 1]:
            ops.append(('equal', a[ii - 1])); ii -= 1; jj -= 1
        elif dp[ii - 1][jj] >= dp[ii][jj - 1]:
            ops.append(('delete', a[ii - 1])); ii -= 1
        else:
            ops.append(('insert', b[jj - 1])); jj -= 1
    while ii > 0:
        ops.append(('delete', a[ii - 1])); ii -= 1
    while jj > 0:
        ops.append(('insert', b[jj - 1])); jj -= 1
    ops.reverse()
    merged = []
    for t, ch in ops:
        if merged and merged[-1][0] == t:
            merged[-1][1] += ch
        else:
            merged.append([t, ch])
    return merged


def change_ratio(ops):
    if not ops:
        return 0.0
    dele = sum(len(t) for op, t in ops if op == 'delete')
    ins = sum(len(t) for op, t in ops if op == 'insert')
    eq = sum(len(t) for op, t in ops if op == 'equal')
    orig = eq + dele
    return (1.0 if ins else 0.0) if orig == 0 else (dele + ins) / orig


def make_el(tag, **attrs):
    el = _make(f'{W}{tag}')
    for k, v in attrs.items():
        el.set(f'{W}{k}', str(v))
    return el


from lxml import etree
_make = etree.Element


def next_rev_id(doc):
    mx = 0
    for tag in ('ins', 'del'):
        for el in doc.element.body.iter(f'{W}{tag}'):
            try:
                mx = max(mx, int(el.get(f'{W}id', 0)))
            except (TypeError, ValueError):
                pass
    return mx + 1


def collect_chars(p):
    """返回 [(char, rPr元素或None的deepcopy)]，只统计 w:t 文本。"""
    chars = []
    for r in p.findall(f'{W}r'):
        rpr = r.find(f'{W}rPr')
        rpr_copy = copy.deepcopy(rpr) if rpr is not None else None
        for t in r.findall(f'{W}t'):
            txt = t.text or ''
            for ch in txt:
                chars.append((ch, rpr_copy))
    return chars


def make_run(text, rpr, deleted=False):
    r = _make(f'{W}r')
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = _make(f'{W}delText' if deleted else f'{W}t')
    t.set(XML_SPACE, 'preserve')
    t.text = text
    r.append(t)
    return r


def rewrite_paragraph(p, ops, chars, rev):
    """按 diff 块重建段落内容：equal 原样保留、delete 包 w:del、insert 包 w:ins。"""
    ppr = p.find(f'{W}pPr')
    for child in list(p):
        if child is not ppr:
            p.remove(child)

    ci = 0  # chars 游标（equal/delete 消耗原文字符）
    prev_rpr = chars[0][1] if chars else None
    for op, text in ops:
        if op == 'equal':
            # 按相邻同格式切 run，尽量保留原格式
            i = 0
            while i < len(text):
                rpr = chars[ci][1]
                j = 0
                while i + j < len(text) and chars[ci + j][1] is rpr:
                    j += 1
                p.append(make_run(text[i:i + j], rpr))
                ci += j
                prev_rpr = rpr
                i += j
        elif op == 'delete':
            ins = make_el('del', id=rev, author=AUTHOR, date=rev_date); rev += 1
            i = 0
            while i < len(text):
                rpr = chars[ci][1]
                j = 0
                while i + j < len(text) and chars[ci + j][1] is rpr:
                    j += 1
                ins.append(make_run(text[i:i + j], rpr, deleted=True))
                ci += j
                i += j
            p.append(ins)
        else:  # insert
            w_ins = make_el('ins', id=rev, author=AUTHOR, date=rev_date); rev += 1
            w_ins.append(make_run(text, prev_rpr))
            p.append(w_ins)
    return rev


def add_flag_comment(doc, p, text):
    """在段落 p 上锚定一条 Word 批注（python-docx 只建内容，锚点需手写
    commentRangeStart/End + commentReference）。返回 True/False。"""
    cm = doc.comments.add_comment(text, author=AUTHOR, initials='AI')
    cid = str(cm.comment_id)
    start = _make(f'{W}commentRangeStart'); start.set(f'{W}id', cid)
    end = _make(f'{W}commentRangeEnd'); end.set(f'{W}id', cid)
    ref_r = _make(f'{W}r')
    ref = _make(f'{W}commentReference'); ref.set(f'{W}id', cid)
    ref_r.append(ref)
    ppr = p.find(f'{W}pPr')
    p.insert(list(p).index(ppr) + 1 if ppr is not None else 0, start)
    p.append(end); p.append(ref_r)
    return True


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    doc_path = Path(sys.argv[1]).resolve()
    rewrites_path = Path(sys.argv[2]).resolve()
    mode = sys.argv[3] if len(sys.argv) > 3 else 'polish'
    if mode not in MODE_CEILING:
        print(f'模式必须是 polish|logic，收到: {mode}'); sys.exit(1)
    out_path = (Path(sys.argv[4]).resolve() if len(sys.argv) > 4
                else doc_path.with_name(doc_path.stem + '-修订.docx'))
    threshold = min(MODE_CEILING[mode], HARD_CEILING)

    rewrites = json.loads(rewrites_path.read_text(encoding='utf-8'))
    rewrites.sort(key=lambda r: r['paraIndex'])

    shutil.copyfile(doc_path, out_path)
    doc = Document(str(out_path))
    paras = [c for c in doc.element.body if c.tag == f'{W}p']

    global rev_date
    rev_date = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    rev = next_rev_id(doc)

    written = skipped = 0
    skip_reasons = []
    for r in rewrites:
        idx, original, new_text = r['paraIndex'], r['original'], str(r['newText'])
        if not new_text.strip():
            skipped += 1; skip_reasons.append(f"id{r['id']} 改写为空"); continue
        if re.search(r'[\r\n]', new_text.strip()):
            skipped += 1; skip_reasons.append(f"id{r['id']} 含多段,跳过以免破坏段落结构"); continue
        if idx < 0 or idx >= len(paras):
            skipped += 1; skip_reasons.append(f"id{r['id']} paraIndex越界"); continue
        p = paras[idx]
        cur_text = ''.join(t.text or '' for t in p.findall(f'.//{W}t')).strip()
        if cur_text != original.strip():
            skipped += 1; skip_reasons.append(f"id{r['id']} 段落内容已变化,跳过"); continue

        chars = collect_chars(p)
        full_text = ''.join(c for c, _ in chars)
        if full_text.strip() != original.strip():
            skipped += 1; skip_reasons.append(f"id{r['id']} 段落含特殊内容,跳过"); continue

        ai = new_text.strip()
        ops = char_diff(full_text, ai)
        if ops is None:
            skipped += 1; skip_reasons.append(f"id{r['id']} 段落过长,diff跳过"); continue
        if all(op == 'equal' for op, _ in ops):
            skipped += 1; skip_reasons.append(f"id{r['id']} 无差异"); continue
        ratio = change_ratio(ops)
        if ratio > threshold:
            skipped += 1; skip_reasons.append(f"id{r['id']} 改动过大({ratio:.0%}),跳过"); continue

        try:
            rev = rewrite_paragraph(p, ops, chars, rev)
            written += 1
        except Exception as e:
            skipped += 1; skip_reasons.append(f"id{r['id']} 写入失败:{e}")

    # 第二遍：挂疑点批注（必须在段落重建之后，否则锚点会被 rewrite_paragraph 清掉；
    # 无差异/被跳过的段落同样要挂——疑点与是否改写无关）
    flagged = 0
    for r in rewrites:
        flag = str(r.get('flag') or '').strip()
        if not flag:
            continue
        idx = r['paraIndex']
        if not (0 <= idx < len(paras)):
            continue
        try:
            add_flag_comment(doc, paras[idx], flag)
            flagged += 1
        except Exception:
            skip_reasons.append(f"id{r['id']} 批注写入失败")

    doc.save(str(out_path))
    summary = {'mode': mode, 'threshold': threshold, 'total': len(rewrites),
               'written': written, 'skipped': skipped, 'flagged': flagged, 'skipReasons': skip_reasons,
               'output': str(out_path)}
    summary_path = out_path.with_name(out_path.stem + '.summary.json')
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding='utf-8')
    msg = f'共 {len(rewrites)} 段：写入修订 {written} 段，跳过 {skipped} 段，批注 {flagged} 条。输出: {out_path}'
    if skip_reasons:
        from collections import Counter
        grouped = Counter(re.sub(r'id\d+ ', '', s) for s in skip_reasons)
        msg += '\n跳过原因: ' + '; '.join(f'{k}×{v}' if v > 1 else k for k, v in grouped.items())
    print(msg)
    print(f'summary: {summary_path}')


if __name__ == '__main__':
    main()
