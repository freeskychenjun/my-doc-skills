# -*- coding: utf-8 -*-
"""分片产物完整性自检：在合并之前快速发现"某片没写出来/写坏了"。

限流（429/1302）或子代理中断时，症状往往是某个 part 文件缺失或半截；
本脚本把它变成一次确定性检查，避免到合并阶段才发现要重跑。

用法:
  python check_parts.py <paragraphs.json> "<part文件模板>" --count 6 [选项]
  例: python check_parts.py paragraphs.json "pass1_part{}.json" --count 6
      python check_parts.py paragraphs.json "rewrites_part{}.json" --count 6
      （模板里 {} 会被替换为 1..count；相对路径按 paragraphs.json 所在目录解析，
        也可用 --part-dir 指定分片目录）

支持两种输出契约（每片自动识别，同一批可混用）:
  A. 全量契约（旧格式，向后兼容）: JSON 数组，含全部分片段落，
     每条 {id, paraIndex, original, newText[, flag]}
  B. 瘦身契约（P2 实测 Arm B）: JSON 对象 {"_meta":{...},"changes":[...]}，
     changes 只含"改动段 / 有 flag 的段"，改动段 {id,newText}、纯 flag 段 {id,flag}，
     不输出 original / paraIndex。

检查项:
  [两种契约共同]
  1. 文件存在、JSON 可解析、非空；
  2. 分片内 id 不重复；id 必须属于本分片（若同目录存在 paragraphs_partN.json）
     或至少属于 paragraphs.json 的 id 全集；跨分片重复 id 硬失败；
  3. newText 若存在：非空串、不含换行；
  4. 覆盖：全量契约要求全集覆盖；瘦身契约改用 _meta 自报数 + lastId 哨兵，
     不要求 id 全集（这正是瘦身的代价，靠哨兵兜）；
  [全量契约]
  5. 每条 original 与 paragraphs.json 的 text 逐字一致（硬失败）；
     newText 必填（缺失/空 → 硬失败）；
  [瘦身契约]
  6. _meta 校验（硬失败）：缺 _meta / part 号不符 / examined ≠ 本片段落数 /
     缺 lastId 或 lastId ≠ 本片最后一段 id（只填 examined 无法证明片尾被读过）；
  7. _meta.changed 与实际"与原文有差异"条数不一致 → 告警（容忍 ±1）；
  8. 条目必须至少带 newText 或 flag 之一（两者都缺 → 硬失败）；
  [三条机械护栏 —— 对两种契约的 newText 同样生效]
  9. 串段守卫（硬失败）：newText 与本段原文相似度 < 0.5 时，若它与**其他段原文**
     相似度 ≥ 0.9 → 判"疑似串段"，整片重跑。文档内重复段落（两段原文本身近似）
     不算串段（实测误报源，必须先排除）。
  10. 上标守卫（硬失败）：newText 的 ²/³ 个数不得多于原文
      （原文未用上标时改写引入上标会破坏全篇体例）。
  11. 数字守卫（告警）：两侧先把 ²/³ 归一成数字再比数字多重集；纯数字
      \\d+(?:\\.\\d+)?，不用连字符连接（历史上误杀过合法修改）。
      "删除重复年份"这类合法修改会改多重集，故只告警不失败。

退出码: 0 全部通过；1 存在问题（控制台列出可重跑的分片）。
--json <out> 写机器可读摘要（供 preflight / postapply 串联）。
--warn-only 把全部硬失败降级为告警（仅用于诊断，不用于放行）。
"""
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

NUM = re.compile(r'\d+(?:\.\d+)?')
SUPS = ('²', '³')
SS_RATIO_LOW = 0.5      # 与本段原文相似度低于此值 ⇒ 存疑
SS_RATIO_HIGH = 0.9     # 与其他段原文相似度高于此值 ⇒ 判串段
DUP_PARA_RATIO = 0.9    # 两段原文互相近似 ⇒ 文档内重复段落，串段守卫须排除


def numbag(text):
    """数字多重集：先把 ²/³ 归一成数字，再取纯数字 token（不用连字符连接）。"""
    t = text
    for sup, plain in (('²', '2'), ('³', '3')):
        t = t.replace(sup, plain)
    return Counter(NUM.findall(t))


def supcount(text):
    return sum(text.count(c) for c in SUPS)


def sim(a, b):
    import difflib
    return difflib.SequenceMatcher(None, a, b).ratio()


def load_json(path):
    with open(path, encoding='utf-8-sig') as f:
        return json.load(f)


class Checker:
    def __init__(self, warn_only=False):
        self.warn_only = warn_only
        self.bad = []          # [(part, problems)]  硬失败
        self.warn = []         # [(part, warnings)]
        self.notes = []

    def add(self, part, problems, warnings=None):
        if problems:
            self.bad.append((part, problems))
        if warnings:
            self.warn.append((part, warnings))

    def flush_part(self, part, hard, soft, ok_msg):
        self.add(part, hard, soft)
        if hard:
            print(f'✗ part{part}: ' + '; '.join(hard[:4]))
            for w in soft[:6]:
                print(f'   ! part{part}: {w}')
        else:
            print(ok_msg)
            for w in soft[:8]:
                print(f'   ! part{part}: {w}')


def guard_entries(part, entries, by_id, hard, soft, all_ids=None):
    """三条机械护栏：串段 / 上标 / 数字。entries 是 [(id, original, newText)]。"""
    by_id = by_id or {}
    # 其他段候选池（原文）；排除 id 自身
    for pid, orig, nt in entries:
        if not isinstance(nt, str) or nt == '' or orig is None:
            continue
        self_ratio = sim(orig, nt)
        # ---- 9. 串段守卫 ----
        if self_ratio < SS_RATIO_LOW:
            best, best_j, best_src = 0.0, None, ''
            for j, otext in by_id.items():
                if j == pid or not isinstance(otext, str) or not otext:
                    continue
                # 文档内重复段落：j 的原文与本段原文本身近似 ⇒ 该候选不作串段证据
                if sim(orig, otext) >= DUP_PARA_RATIO:
                    continue
                r = sim(nt, otext)
                if r > best:
                    best, best_j, best_src = r, j, '原文'
            for oid, _, ont in entries:
                if oid == pid or not isinstance(ont, str) or not ont:
                    continue
                if oid not in by_id or sim(orig, by_id[oid]) >= DUP_PARA_RATIO:
                    continue
                r = sim(nt, ont)
                if r > best:
                    best, best_j, best_src = r, oid, 'newText'
            if best_j is not None and best >= SS_RATIO_HIGH:
                hard.append(f'疑似串段: id{pid} 的 newText 与本段原文相似度仅 '
                            f'{self_ratio:.2f}，却与 id{best_j} 的{best_src}相似度 '
                            f'{best:.2f}（整片重跑）')
        # ---- 10. 上标守卫 ----
        for sup in SUPS:
            if nt.count(sup) > orig.count(sup):
                hard.append(f'id{pid} 引入上标 {sup}（原文 {orig.count(sup)} 处 → '
                            f'改写 {nt.count(sup)} 处）：原文未用上标时不得改写为上标')
        # ---- 11. 数字守卫（只告警）----
        if numbag(orig) != numbag(nt):
            soft.append(f'id{pid} 数字多重集变化（告警）: '
                        f'{sorted(numbag(orig).elements())} → {sorted(numbag(nt).elements())}')


def check_full_part(part, data, expected_ids, by_id, hard, soft, seen):
    """全量契约：id 全集 + original 逐字 + newText 必填。"""
    entries = []
    ids = []
    for r in data:
        pid = r.get('id')
        ids.append(pid)
        if pid not in by_id:
            hard.append(f'id{pid} 不在 paragraphs.json 中')
            continue
        if r.get('original') != by_id[pid]:
            hard.append(f'id{pid} original 与提取文本不一致')
        nt = r.get('newText')
        if not isinstance(nt, str) or not nt.strip():
            hard.append(f'id{pid} newText 为空')
            continue
        if re.search(r'[\r\n]', nt):
            hard.append(f'id{pid} newText 含换行')
        entries.append((pid, by_id[pid], nt))
    if len(ids) != len(set(ids)):
        hard.append('存在重复 id')
    for pid in ids:
        seen.setdefault(pid, []).append(part)
    if expected_ids is not None:
        miss = set(expected_ids) - set(ids)
        extra = set(ids) - set(expected_ids)
        if miss:
            hard.append(f'缺 {len(miss)} 条: {sorted(miss)[:8]}')
        if extra:
            hard.append(f'多 {len(extra)} 条: {sorted(extra)[:8]}')
    guard_entries(part, entries, by_id, hard, soft)
    return len(data)


def check_lean_part(part, data, expected_ids, by_id, hard, soft, seen):
    """瘦身契约：_meta 哨兵 + changes 子集 + 三条护栏。"""
    meta = data.get('_meta')
    changes = data.get('changes')
    if not isinstance(meta, dict):
        hard.append('缺 _meta 自检头（瘦身契约必须带 _meta；若为旧全量格式请勿用对象包裹）')
        meta = {}
    if not isinstance(changes, list):
        hard.append('缺 changes 数组或类型错误')
        return 0
    # ---- _meta.part ----
    if 'part' in meta and meta.get('part') != part:
        hard.append(f'_meta.part={meta.get("part")} 与本分片号 {part} 不符')
    # ---- _meta.examined：本片段落总数 ----
    if expected_ids is None:
        soft.append('未找到 paragraphs_partN.json，无法校验 _meta.examined')
    else:
        ex = meta.get('examined')
        if ex != len(expected_ids):
            hard.append(f'_meta.examined={ex} ≠ 本片段落数 {len(expected_ids)}'
                        f'（疑整片被截断/伪装成无改动，整片重跑）')
        last = meta.get('lastId')
        if last is None:
            hard.append('_meta 缺 lastId 哨兵（瘦身契约必填：本片最后一段 id）；'
                        '只填 examined 时无法证明片尾被读过，整片重跑')
        elif expected_ids and last != expected_ids[-1]:
            hard.append(f'_meta.lastId={last} ≠ 本片最后一段 id {expected_ids[-1]}')

    entries = []
    ids = []
    n_changed = 0
    for r in changes:
        pid = r.get('id')
        ids.append(pid)
        if pid not in by_id:
            hard.append(f'id{pid} 不在 paragraphs.json 中')
            continue
        if expected_ids is not None and pid not in set(expected_ids):
            hard.append(f'id{pid} 不属于本分片')
        has_nt = 'newText' in r
        has_flag = bool(str(r.get('flag') or '').strip())
        if not has_nt and not has_flag:
            hard.append(f'id{pid} 既无 newText 也无 flag（不应出现在 changes 里）')
            continue
        if not has_nt:
            continue                      # 纯 flag 段：合法
        nt = r.get('newText')
        if not isinstance(nt, str):
            hard.append(f'id{pid} newText 类型错误')
            continue
        if nt == '':
            hard.append(f'id{pid} newText 为空串（无差异时应省略该字段，只写 flag）')
            continue
        if re.search(r'[\r\n]', nt):
            hard.append(f'id{pid} newText 含换行')
        if nt != by_id[pid]:
            n_changed += 1
        entries.append((pid, by_id[pid], nt))
    if len(ids) != len(set(ids)):
        hard.append('本片存在重复 id')
    for pid in ids:
        seen.setdefault(pid, []).append(part)
    # ---- _meta.changed ----
    if 'changed' in meta or meta:
        ch = meta.get('changed')
        if ch is not None and ch != n_changed:
            soft.append(f'_meta.changed={ch} ≠ 实际与原文有差异条数 {n_changed}'
                        f'（容忍 ±1，仅告警）')
    guard_entries(part, entries, by_id, hard, soft)
    return len(changes)


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('paragraphs')
    ap.add_argument('template')
    ap.add_argument('--count', type=int, required=True)
    ap.add_argument('--part-dir', default='', help='分片目录（默认与 paragraphs.json 同目录）')
    ap.add_argument('--json', default='', help='写机器可读摘要到该路径')
    ap.add_argument('--warn-only', action='store_true', help='全部硬失败降级为告警（仅诊断用）')
    args = ap.parse_args()

    pdir = Path(args.paragraphs).resolve().parent
    part_dir = Path(args.part_dir).resolve() if args.part_dir else pdir
    paras = load_json(args.paragraphs)
    by_id = {p['id']: p['text'] for p in paras}
    expected_parts = {}
    for i in range(1, args.count + 1):
        f = part_dir / f'paragraphs_part{i}.json'
        if f.exists():
            expected_parts[i] = [p['id'] for p in load_json(f)]

    ck = Checker(warn_only=args.warn_only)
    seen = {}
    modes = {}
    for i in range(1, args.count + 1):
        f = part_dir / args.template.format(i)
        hard, soft = [], []
        data = None
        if not f.exists():
            hard.append('文件缺失')
        else:
            try:
                data = load_json(f)
            except Exception as e:
                hard.append(f'JSON 解析失败: {e}')
        exp = expected_parts.get(i)
        n = 0
        if data is not None:
            if isinstance(data, dict):
                modes[i] = 'lean'
                if not data.get('changes') and not data.get('_meta'):
                    hard.append('JSON 为空或结构不可识别'
                                '（应为 {"_meta":{...},"changes":[...]} 或旧全量数组）')
                else:
                    n = check_lean_part(i, data, exp, by_id, hard, soft, seen)
            elif isinstance(data, list):
                modes[i] = 'full'
                if not data:
                    hard.append('JSON 为空或非数组')
                else:
                    n = check_full_part(i, data, exp, by_id, hard, soft, seen)
            else:
                hard.append('顶层既不是数组也不是对象')
        ck.flush_part(i, hard, soft,
                      f'✓ part{i}: {n} 条（{"瘦身" if modes.get(i) == "lean" else "全量"}）')

    cross = {k: v for k, v in seen.items() if len(v) > 1}
    if cross:
        print(f'✗ 跨分片重复 id: {list(cross.items())[:5]}')
        ck.bad.append(('cross', ['跨分片重复 id']))

    # 覆盖：仅全量契约要求（瘦身契约靠 _meta 哨兵）
    lean_parts = [i for i, m in modes.items() if m == 'lean']
    if expected_parts and not lean_parts:
        covered = set(seen)
        missing_all = set(by_id) - covered
        if missing_all:
            print(f'✗ 未被任何分片覆盖的 id: {len(missing_all)} 个 {sorted(missing_all)[:10]}')
            ck.bad.append(('coverage', [f'缺 {len(missing_all)} 条']))
        else:
            print(f'✓ 覆盖完整: {len(by_id)} 条全部有产出')
    elif lean_parts:
        print(f'· 瘦身契约分片 {sorted(lean_parts)}：不校验 id 全集，改由 _meta.examined+lastId 兜底')

    n_soft = sum(len(w) for _, w in ck.warn)
    hard_bad = ck.bad if not args.warn_only else []
    if args.json:
        summary = {
            'ok': not hard_bad,
            'modes': modes,
            'hard_failures': [{'part': b[0], 'problems': b[1]} for b in ck.bad],
            'warnings': [{'part': w[0], 'problems': w[1]} for w in ck.warn],
            'parts': sorted(m for m in modes),
            'count': args.count,
        }
        Path(args.json).write_text(json.dumps(summary, ensure_ascii=False, indent=1),
                                  encoding='utf-8')
        print(f'摘要 → {args.json}')

    if hard_bad:
        print(f'\n结论: 有问题 {len(hard_bad)} 处 → 需要重跑的分片: '
              f'{[b[0] for b in hard_bad if isinstance(b[0], int)]}')
        sys.exit(1)
    if n_soft:
        print(f'\n结论: 无硬失败，{n_soft} 条告警（数字多重集/自报数），可进入合并')
    else:
        print('\n结论: 全部分片完好，可以进入合并')


if __name__ == '__main__':
    main()
