# -*- coding: utf-8 -*-
"""apply_verdicts.py —— 把查证遍裁决落回 merged.json，并自动产出 finalize 配置。

背景：第 6 步查证遍的裁决（fix / keep / upgrade_flag）以前要手写脚本合并，而且最容易漏的一件事是
**「fix 改了数字（例如改正法规/标准编号）的段必须登记进 finalize_config.json 的 num_exempt」**——
不登记就会被收口时的数字 guard 判为"擅自改数"而静默回退。本脚本把这两件事一起做掉。

用法:
  python apply_verdicts.py --workdir <运行目录> [--merged merged.json] [--verdicts verify1.json ...]
      [--out merged.verified.json] [--config finalize_config.json] [--flags flags.json]
      [--report apply-verdicts.json] [--dry] [--strict]

默认行为:
  --merged     <workdir>/merged.json
  --verdicts   自动 glob <workdir>/verify*.json
  --out        <workdir>/merged.verified.json
  --config     <workdir>/finalize_config.json（已存在则**只更新 num_exempt**，其余键原样保留）

裁决语义（与 SKILL.md 第 6 步一致）:
  fix            newText 换成裁决文本；flag 写"已核实并改正：依据…"
                 ↳ 若裁决文本与当前文本相同（说明改写遍已经改好），视为 fix_noop：**丢弃 flag**，不留噪音
  keep           弃改：删掉该段 flag，newText 不变
  upgrade_flag   newText 不变，flag 换成新文案

安全网:
  除裁决 fix 的段之外，若还有段落「数字多重集相对原文发生变化」却没有被登记豁免，脚本会**单独列出**
  （unverified_numeric_changes）——这些段在收口时会被数字 guard 回退，需要父代理明确决定。

退出码: 0 正常（含仅有告警）; 1 致命（裁决缺 newText / action 非法 / id 不在 merged 里，或 --strict 下漏裁）
"""
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

NUM = re.compile(r'\d+(?:\.\d+)?')
SUP = {'²': '2', '³': '3'}
ACTIONS = ('fix', 'keep', 'upgrade_flag')


def numbag(text):
    """数字 token 多重集：先把 ²/³ 归一成普通数字，避免 km2→km² 被误判成数字变化。"""
    t = text or ''
    for s, p in SUP.items():
        t = t.replace(s, p)
    return Counter(NUM.findall(t))


def load_config(path: Path) -> dict:
    base = {'name': '', 'typo': {}, 'term': {}, 'reffix': {}, 'reffix_note': {},
            'num_exempt': [], 'flag_final': {}, 'clear_flags': []}
    if path.exists():
        try:
            old = json.loads(path.read_text(encoding='utf-8'))
            base.update(old)
        except Exception as e:
            print(f'! 已有 {path.name} 读取失败，将覆盖: {e}')
    return base


def main():
    ap = argparse.ArgumentParser(description='查证遍裁决落库 + 自动产出 finalize_config.json')
    ap.add_argument('--workdir', required=True)
    ap.add_argument('--merged', default='')
    ap.add_argument('--verdicts', nargs='*', default=[])
    ap.add_argument('--out', default='')
    ap.add_argument('--config', default='')
    ap.add_argument('--flags', default='', help='collect_flags.py 的 flags.json，用于覆盖率核对')
    ap.add_argument('--report', default='', help='机器可读报告，默认 <workdir>/apply-verdicts.json')
    ap.add_argument('--dry', action='store_true', help='只算不写')
    ap.add_argument('--strict', action='store_true', help='有 flag 未被裁决即失败')
    ap.add_argument('--exempt', default='',
                    help='父代理复核后决定放行的 id（逗号分隔）：这些段允许数字多重集变化，'
                         '写进 finalize_config.json 的 num_exempt。用于处理脚本报出的'
                         'unverified_numeric_changes（如"删去重复年份"这类合法改动）')
    args = ap.parse_args()

    wd = Path(args.workdir).resolve()
    merged_path = Path(args.merged) if args.merged else wd / 'merged.json'
    out_path = Path(args.out) if args.out else wd / 'merged.verified.json'
    cfg_path = Path(args.config) if args.config else wd / 'finalize_config.json'
    rep_path = Path(args.report) if args.report else wd / 'apply-verdicts.json'
    if not merged_path.exists():
        print(f'✗ 找不到 merged: {merged_path}', file=sys.stderr)
        sys.exit(1)

    merged = json.loads(merged_path.read_text(encoding='utf-8'))
    idx = {e['id']: e for e in merged}

    vfiles = [Path(x) for x in args.verdicts] or sorted(wd.glob('verify*.json'))
    vfiles = [p for p in vfiles if p.exists() and p.name != 'apply-verdicts.json']
    if not vfiles:
        print(f'✗ 没找到裁决文件（默认 glob {wd}\\verify*.json）', file=sys.stderr)
        sys.exit(1)

    verdicts, fatal = [], []
    for f in vfiles:
        try:
            d = json.loads(f.read_text(encoding='utf-8'))
        except Exception as e:
            fatal.append(f'{f.name}: JSON 解析失败 {e}')
            continue
        vs = d.get('verdicts') if isinstance(d, dict) else d
        if not isinstance(vs, list):
            fatal.append(f'{f.name}: 没有 verdicts 数组')
            continue
        print(f'  · {f.name}: {len(vs)} 条')
        for v in vs:
            v['_file'] = f.name
            verdicts.append(v)

    stats = Counter()
    fixed, kept, upgraded, noop = [], [], [], []
    num_exempt = []
    for v in verdicts:
        pid, act = v.get('id'), v.get('action')
        if act not in ACTIONS:
            fatal.append(f'{v["_file"]}: id{pid} action 非法: {act!r}')
            continue
        if pid not in idx:
            fatal.append(f'{v["_file"]}: id{pid} 不在 merged.json 中')
            continue
        e = idx[pid]
        before = e.get('newText') if e.get('newText') is not None else e['original']
        if act == 'fix':
            nt = v.get('newText')
            if not nt:
                fatal.append(f'{v["_file"]}: id{pid} action=fix 但缺 newText')
                continue
            if nt == before:
                noop.append(pid)                 # 改写遍已经改好，丢弃 flag，不留噪音
                e.pop('flag', None)
            else:
                e['newText'] = nt
                e['flag'] = v.get('note') or ''
                fixed.append(pid)
                if numbag(e['original']) != numbag(nt):
                    num_exempt.append(pid)
        elif act == 'keep':
            before_flag = e.pop('flag', None)
            if before_flag:
                kept.append(pid)
        else:                                    # upgrade_flag
            e['flag'] = v.get('note') or e.get('flag') or ''
            upgraded.append(pid)
        stats[act] += 1

    for e in merged:                             # 清掉空 flag
        if e.get('flag') is not None and not str(e['flag']).strip():
            e.pop('flag', None)

    # 安全网：非裁决 fix、但数字多重集变了、又没被豁免的段 —— 收口时会被 guard 回退
    fixed_set = set(fixed)
    unverified_num = []
    for e in merged:
        nt, o = e.get('newText'), e['original']
        if nt is None or nt == o or e['id'] in fixed_set:
            continue
        if numbag(nt) != numbag(o):
            unverified_num.append(e['id'])

    # 覆盖率
    total_flags, uncovered = None, []
    flags_path = Path(args.flags) if args.flags else (wd / 'flags.json')
    if flags_path.exists():
        fd = json.loads(flags_path.read_text(encoding='utf-8'))
        items = fd.get('items', []) if isinstance(fd, dict) else fd
        total_flags = len(items)
        done = {v['id'] for v in verdicts}
        uncovered = sorted({it['id'] for it in items} - done)

    # 父代理复核后手动放行的段（工具无法判断"删重复年份"这类改动是否合法）
    manual_exempt = []
    for tok in str(args.exempt).replace('，', ',').split(','):
        tok = tok.strip()
        if not tok:
            continue
        if not tok.isdigit():
            fatal.append(f'--exempt 里的 id 不是整数: {tok!r}')
            continue
        manual_exempt.append(int(tok))
    missing_ex = [x for x in manual_exempt if x not in idx]
    if missing_ex:
        fatal.append(f'--exempt 里这些 id 不在 merged.json 中: {missing_ex}')
    cfg = load_config(cfg_path)
    before_ex = set(cfg.get('num_exempt') or [])
    # 已豁免过的（含本脚本上一次运行写进配置的）不再重复提醒，保证可重复执行
    still_unverified = [x for x in unverified_num if x not in (before_ex | set(manual_exempt))]
    cfg['num_exempt'] = sorted(before_ex | set(num_exempt) | set(manual_exempt))
    if not cfg.get('name'):
        cfg['name'] = wd.name
    flags_left = sum(1 for e in merged if e.get('flag'))
    changed = sum(1 for e in merged if e.get('newText') != e['original'])

    print(f'\n裁决: {dict(stats)}（共 {len(verdicts)} 条，来自 {len(vfiles)} 个文件）')
    print(f'  fix 生效 {len(fixed)} / fix 无操作(改写遍已改好) {len(noop)} / 弃改 {len(kept)} / 升级批注 {len(upgraded)}')
    print(f'  merged → {len(merged)} 段，改动 {changed} 段，批注 {flags_left} 条')
    print(f'  自动登记 num_exempt: {sorted(set(num_exempt))}（原有 {sorted(before_ex)}）')
    if manual_exempt:
        print(f'  父代理手动放行: {sorted(set(manual_exempt))}')
    if still_unverified:
        print(f'  ⚠ 未裁决但数字多重集变化的段（收口时会被数字 guard 回退，请明确决定；'
              f'复核后加 --exempt "{",".join(str(x) for x in still_unverified)}"）：{still_unverified[:12]}')
    if total_flags is not None:
        print(f'  覆盖率: {total_flags - len(uncovered)}/{total_flags}' +
              (f'，漏裁 {uncovered}' if uncovered else '（全部有裁决）'))
    if fatal:
        print('\n✗ 致命问题:')
        for x in fatal:
            print('  - ' + x)
    if args.strict and uncovered:
        fatal.append(f'--strict: {len(uncovered)} 条 flag 未裁决')
        print('✗ --strict：存在未裁决 flag')

    if not args.dry and not fatal:
        out_path.write_text(json.dumps(merged, ensure_ascii=False), encoding='utf-8')
        cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding='utf-8')
        rep_path.write_text(json.dumps({
            'tool': 'apply_verdicts.py', 'merged_in': str(merged_path), 'out': str(out_path),
            'config': str(cfg_path), 'verdict_files': [str(p) for p in vfiles],
            'stats': dict(stats), 'fixed': fixed, 'fix_noop': noop, 'kept': kept,
            'upgraded': upgraded, 'num_exempt_added': sorted(set(num_exempt)),
            'num_exempt_manual': sorted(set(manual_exempt)),
            'unverified_numeric_changes': unverified_num,
            'unverified_remaining': still_unverified,
            'flags_total': total_flags, 'flags_uncovered': uncovered,
            'changed_paragraphs': changed, 'flags_left': flags_left,
        }, ensure_ascii=False, indent=1), encoding='utf-8')
        print(f'\n→ {out_path}')
        print(f'→ {cfg_path}')
        print(f'→ {rep_path}')
        print('下一步: python finalize.py <paragraphs.json> --merged "%s" --config "%s" --dry' % (out_path, cfg_path))
    elif args.dry:
        print('\n(dry run，未写盘)')

    sys.exit(1 if fatal else 0)


if __name__ == '__main__':
    main()
