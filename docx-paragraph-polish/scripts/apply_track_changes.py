# -*- coding: utf-8 -*-
"""以 Word 修订(track changes)模式把逐段改写写回 docx 副本。

纯 OpenXML 实现（python-docx + lxml），不依赖本机安装 Word/WPS；
产物是带 w:ins/w:del 修订标记的标准 docx，Word 与 WPS 均可逐条接受/拒绝。

移植自 WordTools/SmartWriting/{CharDiff,TrackChangesRewriter}.cs：
- 字符级 LCS diff，删字标删除修订、增字标插入修订，未改动部分原样保留；
- 净改动占比超门槛（润色 0.6 / 逻辑重构 0.95，硬顶 0.95）自动跳过；
- 段首 sanity 校验：段落文本与 extract 时不一致则跳过。

本版新增（P4 改造，2026-10-17；核心 diff/写回算法未改）：
1) 段内**附加元素保留**：段落级的 bookmarkStart/End、m:oMath 公式、批注锚点，以及
   含非 w:t 内容的 run（w:tab / w:fldChar / w:instrText / w:drawing / w:noBreakHyphen 等）
   在重建段落时按"原文字符位置→新文字符位置"映射插回原位，不再被静默丢弃
   （原先只能靠每个运行目录现抄一份 apply_tc_safe.py，实测会丢对照表 tab 对齐与书签）；
2) --protect <protected_ids.json|json数组>：保护清单内的段落不写修订（只挂批注）；
3) --numeric-guard：写回前再比一次数字 token 多重集（²/³ 先归一成数字），
   不一致的段落回退并记入 summary（这是"数字红线"的第二道机械闸门，
   与 finalize.py 的 guard 互为兜底；--num-exempt 或条目里 "numExempt": true 可豁免）。

用法:
  python apply_track_changes.py <docx> <rewrites.json> [polish|logic] [输出docx路径] \
      [--protect protected_ids.json] [--numeric-guard] [--num-exempt 50,577] [--no-keep-extras]
rewrites.json: [{"id":1,"paraIndex":3,"original":"...","newText":"..."}, ...]
（paraIndex/original 必须来自 extract_body_paragraphs.py 的输出，不得手改）
"""
import argparse
import copy
import json
import re
import shutil
import sys
import datetime
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
AUTHOR = 'AI校正'
HARD_CEILING = 0.95
MODE_CEILING = {'polish': 0.6, 'logic': 0.95}
MAX_MATRIX = 4_000_000
NUM = re.compile(r'\d+(?:\.\d+)?')
SUPS = (('²', '2'), ('³', '3'))


def numbag(text):
    """数字 token 多重集：先把 ²/³ 归一成数字再比（否则 km2→km² 会被误判成删数字）。
    只取纯数字 \\d+(?:\\.\\d+)?，不用连字符连接（历史上误杀过合法修改）。"""
    t = text
    for sup, plain in SUPS:
        t = t.replace(sup, plain)
    return Counter(NUM.findall(t))


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


def collect_layout(p):
    """返回 (chars, extras)。

    chars: [(char, rPr元素或None的deepcopy)]，只统计"纯文本 run"里的 w:t；
    extras: [(原文字符偏移, deepcopy(元素))]，是段落里**不能重建**的东西：
      - 含非 w:t 子元素的 run（w:tab / w:fldChar / w:instrText / w:drawing /
        w:object / w:noBreakHyphen / w:br / w:sym / OMML m:oMath 的 run 包装等）——
        整体保留，否则重建时会静默丢弃（对照表 tab 对齐、域代码、图片会坏）；
      - 段落级的直接子元素（bookmarkStart/End、m:oMath、commentRangeStart 等）。
    """
    chars, extras = [], []
    for child in list(p):
        if not isinstance(child.tag, str):
            continue
        if child.tag == f'{W}pPr':
            continue
        if child.tag == f'{W}r':
            inner = [k for k in child if k.tag not in (f'{W}rPr', f'{W}t')]
            if inner:
                extras.append((len(chars), copy.deepcopy(child)))
                continue
            rpr = child.find(f'{W}rPr')
            rpr_copy = copy.deepcopy(rpr) if rpr is not None else None
            for t in child.findall(f'{W}t'):
                for ch in (t.text or ''):
                    chars.append((ch, rpr_copy))
        else:
            extras.append((len(chars), copy.deepcopy(child)))
    return chars, extras


def collect_chars(p):
    """兼容旧接口（fix_captions.py 依赖）：只返回 [(char, rPr)]。"""
    return collect_layout(p)[0]


def map_offsets(ops, offsets, new_len):
    """把"原文字符坐标"的偏移沿 diff 映射到"新文字符坐标"。"""
    out = []
    for o in offsets:
        old_acc = new_acc = 0
        mapped = None
        for op, text in ops:
            n = len(text)
            if old_acc == o:
                mapped = new_acc
                break
            if op == 'equal':
                if o < old_acc + n:
                    mapped = new_acc + (o - old_acc)
                    break
                old_acc += n
                new_acc += n
            elif op == 'delete':
                if o < old_acc + n:
                    mapped = new_acc
                    break
                old_acc += n
            else:
                new_acc += n
        if mapped is None:
            mapped = new_acc
        out.append(max(0, min(mapped, new_len)))
    return out


def make_run(text, rpr, deleted=False):
    r = _make(f'{W}r')
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = _make(f'{W}delText' if deleted else f'{W}t')
    t.set(XML_SPACE, 'preserve')
    t.text = text
    r.append(t)
    return r


def rewrite_paragraph(p, ops, chars, rev, extras=None):
    """按 diff 块重建段落内容：equal 原样保留、delete 包 w:del、insert 包 w:ins。

    extras（可选）：[(新文字符偏移, 元素)]，为 None/空时行为与旧版逐字一致；
    给出时把元素按位置插回重建后的段落，避免丢书签/公式/域/tab。
    """
    ppr = p.find(f'{W}pPr')
    for child in list(p):
        if child is not ppr:
            p.remove(child)

    extras = extras or []
    state = {'ci': 0, 'prev_rpr': chars[0][1] if chars else None,
             'ei': 0, 'pos_new': 0}

    def emit_runs(container, text, deleted=False):
        i = 0
        while i < len(text):
            rpr = chars[state['ci']][1]
            j = 0
            while i + j < len(text) and chars[state['ci'] + j][1] is rpr:
                j += 1
            container.append(make_run(text[i:i + j], rpr, deleted))
            state['prev_rpr'] = rpr
            state['ci'] += j
            i += j

    def flush(upto):
        while state['ei'] < len(extras) and extras[state['ei']][0] <= upto:
            p.append(extras[state['ei']][1])
            state['ei'] += 1

    for op, text in ops:
        if op == 'equal':
            if not extras:
                emit_runs(p, text)
                state['pos_new'] += len(text)
                continue
            i = 0
            while i < len(text):
                base = state['pos_new'] + i
                flush(base)
                nxt = len(text)
                if state['ei'] < len(extras):
                    nxt = min(nxt, max(i, extras[state['ei']][0] - state['pos_new']))
                if nxt <= i:
                    flush(base)
                    nxt = len(text)
                emit_runs(p, text[i:nxt])
                i = nxt
            state['pos_new'] += len(text)
        elif op == 'delete':
            flush(state['pos_new'])
            ins = make_el('del', id=rev, author=AUTHOR, date=rev_date); rev += 1
            emit_runs(ins, text, deleted=True)
            p.append(ins)
        else:  # insert
            flush(state['pos_new'])
            w_ins = make_el('ins', id=rev, author=AUTHOR, date=rev_date); rev += 1
            w_ins.append(make_run(text, state['prev_rpr']))
            p.append(w_ins)
            state['pos_new'] += len(text)
    flush(1 << 30)
    return rev


def restore_ws(p, prefix, suffix, core_chars):
    """把段首/段尾原有空白按原格式写回（不产生修订，属排版留白）。"""
    rpr = core_chars[0][1] if core_chars else None
    ppr = p.find(f'{W}pPr')
    pos = list(p).index(ppr) + 1 if ppr is not None else 0
    if prefix:
        p.insert(pos, make_run(prefix, rpr))
    if suffix:
        p.append(make_run(suffix, rpr))


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


def load_protect(path):
    """读保护清单：支持 {id: 原因} 字典（finalize.py 产物）或 [id, ...] 数组。
    返回 {id: 原因}。"""
    if not path:
        return {}
    d = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if isinstance(d, dict):
        return {int(k): str(v) for k, v in d.items()}
    return {int(k): '保护清单' for k in d}


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('docx')
    ap.add_argument('rewrites')
    ap.add_argument('mode', nargs='?', default='polish', choices=list(MODE_CEILING))
    ap.add_argument('out', nargs='?', default='')
    ap.add_argument('--protect', default='',
                    help='保护清单 json（{id:原因} 或 [id,...]）；清单内段落不写修订')
    ap.add_argument('--numeric-guard', action='store_true',
                    help='写回前再比一次数字 token 多重集（²/³ 归一），不一致则回退')
    ap.add_argument('--num-exempt', default='',
                    help='数字 guard 豁免的 id 列表，逗号分隔（如 50,577）')
    ap.add_argument('--no-keep-extras', action='store_true',
                    help='不保留段内附加元素（书签/公式/域/tab）——仅用于复现旧行为')
    args = ap.parse_args()

    doc_path = Path(args.docx).resolve()
    rewrites_path = Path(args.rewrites).resolve()
    mode = args.mode
    out_path = (Path(args.out).resolve() if args.out
                else doc_path.with_name(doc_path.stem + '-修订.docx'))
    threshold = min(MODE_CEILING[mode], HARD_CEILING)
    protect = load_protect(args.protect)
    keep_extras = not args.no_keep_extras
    num_exempt = {int(x) for x in re.split(r'[,，\s]+', args.num_exempt) if x.strip().isdigit()}

    rewrites = json.loads(rewrites_path.read_text(encoding='utf-8'))
    rewrites.sort(key=lambda r: r['paraIndex'])

    shutil.copyfile(doc_path, out_path)
    doc = Document(str(out_path))
    paras = [c for c in doc.element.body if c.tag == f'{W}p']

    global rev_date
    rev_date = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    rev = next_rev_id(doc)

    written = skipped = preserved = protected_n = numguard_n = 0
    skip_reasons = []
    outcome = {}  # id -> (written: bool, reason: str)  供批注一致性防护
    for r in rewrites:
        idx, original, new_text = r['paraIndex'], r['original'], str(r['newText'])
        if r['id'] in protect:
            protected_n += 1
            skip_reasons.append(f"id{r['id']} 保护清单({protect[r['id']]})")
            outcome[r['id']] = (False, '保护清单')
            continue
        if not new_text.strip():
            skipped += 1; skip_reasons.append(f"id{r['id']} 改写为空")
            outcome[r['id']] = (False, '改写为空'); continue
        if re.search(r'[\r\n]', new_text.strip()):
            skipped += 1; skip_reasons.append(f"id{r['id']} 含多段,跳过以免破坏段落结构")
            outcome[r['id']] = (False, '含多段'); continue
        if idx < 0 or idx >= len(paras):
            skipped += 1; skip_reasons.append(f"id{r['id']} paraIndex越界")
            outcome[r['id']] = (False, 'paraIndex越界'); continue
        p = paras[idx]
        cur_text = ''.join(t.text or '' for t in p.findall(f'.//{W}t')).strip()
        if cur_text != original.strip():
            skipped += 1; skip_reasons.append(f"id{r['id']} 段落内容已变化,跳过")
            outcome[r['id']] = (False, '段落内容已变化'); continue

        chars, extras = collect_layout(p)
        full_text = ''.join(c for c, _ in chars)
        if full_text.strip() != original.strip():
            skipped += 1; skip_reasons.append(f"id{r['id']} 段落含特殊内容,跳过")
            outcome[r['id']] = (False, '段落含特殊内容'); continue

        # 首尾空白（缩进/居中留白）原样保留，只在核心文本上做 diff：
        # 否则 strip 后的 newText 与未 strip 的原文相比，会把段首空白当成"删除"写进修订。
        core = full_text.strip()
        prefix = full_text[:len(full_text) - len(full_text.lstrip())]
        suffix = full_text[len(full_text.rstrip()):]
        core_chars = chars[len(prefix):len(chars) - len(suffix)] if suffix else chars[len(prefix):]

        ai = new_text.strip()
        if core == ai:
            skipped += 1; skip_reasons.append(f"id{r['id']} 无差异")
            outcome[r['id']] = (False, '无差异'); continue
        if args.numeric_guard and r['id'] not in num_exempt and not r.get('numExempt'):
            if numbag(core) != numbag(ai):
                numguard_n += 1
                skipped += 1
                skip_reasons.append(f"id{r['id']} 数字guard回退(多重集不一致)")
                outcome[r['id']] = (False, '数字guard回退'); continue
        ops = char_diff(core, ai)
        if ops is None:
            skipped += 1; skip_reasons.append(f"id{r['id']} 段落过长,diff跳过")
            outcome[r['id']] = (False, '段落过长'); continue
        if all(op == 'equal' for op, _ in ops):
            skipped += 1; skip_reasons.append(f"id{r['id']} 无差异")
            outcome[r['id']] = (False, '无差异'); continue
        ratio = change_ratio(ops)
        if ratio > threshold:
            # 净改动占比超门槛时再用序列相似度复核：短段的 1:1 字符替换（如
            # "闸门名称"→"泵站名称"）净改动可达 100%，却仍是纯纠错，不该被误判跳过
            # （泸水第十二例实测形态）。相似度也低才认定"改动过大"。
            import difflib as _dl
            sim = _dl.SequenceMatcher(None, core, ai).ratio()
            if sim < 1 - threshold:
                skipped += 1; skip_reasons.append(f"id{r['id']} 改动过大({ratio:.0%},相似度{sim:.0%}),跳过")
                outcome[r['id']] = (False, f'改动过大({ratio:.0%})'); continue

        try:
            if keep_extras and extras:
                core_offsets = [max(0, min(off - len(prefix), len(core))) for off, _ in extras]
                mapped = map_offsets(ops, core_offsets, len(ai))
                new_extras = [(mapped[k], extras[k][1]) for k in range(len(extras))]
                rev = rewrite_paragraph(p, ops, core_chars, rev, new_extras)
                preserved += 1
            else:
                rev = rewrite_paragraph(p, ops, core_chars, rev)
            if prefix or suffix:
                restore_ws(p, prefix, suffix, core_chars)
            written += 1
            outcome[r['id']] = (True, '')
        except Exception as e:
            skipped += 1; skip_reasons.append(f"id{r['id']} 写入失败:{e}")
            outcome[r['id']] = (False, f'写入失败:{e}')

    # 第二遍：挂疑点批注（必须在段落重建之后，否则锚点会被 rewrite_paragraph 清掉；
    # 无差异/被跳过的段落同样要挂——疑点与是否改写无关）
    flagged = 0
    stale_flagged = []  # 批注声称"已改正"但修订未写入的段（批注-修订一致性空洞）
    for r in rewrites:
        flag = str(r.get('flag') or '').strip()
        if not flag:
            continue
        idx = r['paraIndex']
        if not (0 <= idx < len(paras)):
            continue
        ok, reason = outcome.get(r['id'], (True, ''))
        if not ok and reason != '无差异':
            # 该段修订未写入，但 flag 若宣称"已改正/已替换/已改为"则与事实不符——加警示前缀
            if re.search(r'已核实并改正|已核实改正|已替换|已改为|已改正|已删去|已补全', flag):
                flag = ('【警示：本段修订未写入（原因：%s），下述"已改正"内容实际未生效，请手动处理！】' % reason) + flag
                stale_flagged.append(f"id{r['id']}({reason})")
        try:
            add_flag_comment(doc, paras[idx], flag)
            flagged += 1
        except Exception:
            skip_reasons.append(f"id{r['id']} 批注写入失败")
    if stale_flagged:
        skip_reasons.append('批注声称已改但修订未写入: ' + ', '.join(stale_flagged))

    doc.save(str(out_path))
    summary = {'mode': mode, 'threshold': threshold, 'total': len(rewrites),
               'written': written, 'skipped': skipped, 'flagged': flagged,
               'protected': protected_n, 'numeric_guard_skipped': numguard_n,
               'paragraphs_with_preserved_extras': preserved,
               'skipReasons': skip_reasons, 'output': str(out_path)}
    summary_path = out_path.with_name(out_path.stem + '.summary.json')
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding='utf-8')
    msg = (f'共 {len(rewrites)} 段：写入修订 {written} 段，跳过 {skipped} 段，批注 {flagged} 条。')
    if protected_n or numguard_n or preserved:
        msg += (f'\n保护清单跳过 {protected_n} 段；数字 guard 回退 {numguard_n} 段；'
                f'保留段内附加元素（书签/公式/域/tab）{preserved} 段。')
    msg += f'\n输出: {out_path}'
    if skip_reasons:
        from collections import Counter
        grouped = Counter(re.sub(r'id\d+ ', '', s) for s in skip_reasons)
        msg += '\n跳过原因: ' + '; '.join(f'{k}×{v}' if v > 1 else k for k, v in grouped.items())
    print(msg)
    print(f'summary: {summary_path}')


if __name__ == '__main__':
    main()
