# -*- coding: utf-8 -*-
"""多遍改写结果合并：lint 自动修复 + 投票 + 冲突/低置信清单 + 仲裁。

用法:
  python merge_passes.py paragraphs.json -o merged.json \
      --passes pass1.json pass2.json \
      [--lint paragraphs-lint.json] \
      [--arb rewrites_arb.json] [--adopt-lowconf]

输入:
- paragraphs.json      extract 输出（id/paraIndex/text）
- --passes             N 遍改写的 rewrites json（≥2 遍）
- --lint               lint_paragraphs.py 输出（可选；auto_fix 不依赖它，只用于统计）
- --arb                仲裁结果 [{id,newText,flag?}]，解决 conflicts/lowconf
- --adopt-lowconf      无仲裁时直接采纳单票改动（默认不采纳，保持原文）

合并策略:
1. ≥2 遍改出完全相同的文本 → 采纳（高置信）。
2. 多遍都改但文本互不相同 → 冲突，等仲裁；无仲裁则保持原文并列入报告。
3. 仅 1 遍改 → 低置信；无仲裁默认保持原文（--adopt-lowconf 可改为采纳）。
4. 所有段落最后叠加 auto_fix 零风险修复（错别字映射/单位大小写/空格）。
5. flag 字段：各遍/仲裁提供的疑点全部保留（写回时转 Word 批注）。

输出: merged.json（rewrites 格式，可直接给 apply_track_changes.py）+ 控制台报告。
"""
import argparse
import json
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from lint_paragraphs import auto_fix

sys.stdout.reconfigure(encoding='utf-8', errors='replace')


def unify_unit_style(original, text):
    """单位体例跟随原文：原文未用上标（²/³）时，把改写文本里的上标还原为普通数字。
    避免局部段落用 m³/km²、其余数百处仍用 m3/km2 造成的体例不一致（对审后修编稿
    尤其重要——体例统一属于全篇决策，不该由逐段改写顺手决定）。"""
    changed = []
    for sup, plain in (('²', '2'), ('³', '3')):
        if sup not in original and sup in text:
            text = text.replace(sup, plain)
            changed.append(f'{sup}→{plain}')
    return text, changed


def dedupe_flags(flags):
    """两遍可能提同一条疑点（措辞略有差异），保留信息量更大的一条。"""
    import difflib
    out = []
    for f in flags:
        dup = None
        for i, kept in enumerate(out):
            if f in kept or kept in f:
                dup = i
                break
            if difflib.SequenceMatcher(None, f, kept).ratio() > 0.5:
                dup = i
                break
        if dup is None:
            out.append(f)
        elif len(f) > len(out[dup]):
            out[dup] = f
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('paragraphs')
    ap.add_argument('-o', '--out', required=True)
    ap.add_argument('--passes', nargs='+', required=True)
    ap.add_argument('--lint', default='')
    ap.add_argument('--arb', default='')
    ap.add_argument('--adopt-lowconf', action='store_true')
    ap.add_argument('--no-unit-consistency', action='store_true',
                    help='关闭"单位体例跟随原文"（默认开启：原文无上标则还原 m³→m3、km²→km2）')
    args = ap.parse_args()

    paras = json.load(open(args.paragraphs, encoding='utf-8'))
    passes = [json.load(open(f, encoding='utf-8-sig')) for f in args.passes]
    arb = {r['id']: r for r in json.load(open(args.arb, encoding='utf-8-sig'))} if args.arb else {}

    stats = Counter()
    report = {'adopted_2plus': [], 'conflict': [], 'lowconf': [], 'arb_resolved': [],
              'autofix_only': [], 'unchanged': []}
    merged = []
    for p in paras:
        i, orig = p['id'], p['text']
        texts = []
        flags = []
        for pw in passes:
            r = next((x for x in pw if x['id'] == i), None)
            if r:
                texts.append(r['newText'])
                if r.get('flag'):
                    flags.append(r['flag'])
        a = arb.get(i)
        if a:
            if a.get('flag'):
                flags.append(a['flag'])
            flags = list(dict.fromkeys(flags))

        changed = [t for t in texts if t != orig]
        votes = Counter(changed)
        top_text, top_n = (votes.most_common(1)[0] if votes else (None, 0))

        if i in arb:
            chosen, why = arb[i]['newText'], 'arb'
            stats['仲裁采纳'] += 1
            report['arb_resolved'].append(i)
        elif top_n >= 2:
            chosen, why = top_text, 'vote'
            stats['≥2票一致'] += 1
            report['adopted_2plus'].append(i)
        elif len(changed) >= 2:
            chosen, why = orig, 'conflict'
            stats['冲突待仲裁'] += 1
            report['conflict'].append({'id': i, 'candidates': list(dict.fromkeys(changed))})
        elif len(changed) == 1:
            if args.adopt_lowconf:
                chosen, why = changed[0], 'lowconf-adopt'
                stats['单票采纳'] += 1
            else:
                chosen, why = orig, 'lowconf-keep'
                stats['单票保持原文'] += 1
            report['lowconf'].append({'id': i, 'candidate': changed[0]})
        else:
            chosen, why = orig, 'keep'

        fixed, applied = auto_fix(chosen)
        if fixed != orig and why in ('keep', 'lowconf-keep'):
            stats['仅lint自动修复'] += 1
            report['autofix_only'].append({'id': i, 'rules': applied})
        if not args.no_unit_consistency:
            fixed2, sup_changed = unify_unit_style(orig, fixed)
            if sup_changed:
                stats['单位体例还原'] += 1
                report.setdefault('unit_style_reverted', []).append({'id': i, 'rules': sup_changed})
                fixed = fixed2

        entry = {'id': i, 'paraIndex': p['paraIndex'], 'original': orig, 'newText': fixed}
        if flags:
            entry['flag'] = '；'.join(dedupe_flags(list(dict.fromkeys(flags))))
        merged.append(entry)

    json.dump(merged, open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)
    n_changed = sum(1 for m in merged if m['newText'] != m['original'])
    print(f'合并完成：{len(merged)} 段，最终有改动 {n_changed} 段 → {args.out}')
    for k, v in stats.most_common():
        print(f'  {k}: {v}')
    if report['conflict']:
        print(f"冲突待仲裁段: {[c['id'] for c in report['conflict']]}")
    rp = Path(args.out).with_suffix('.merge-report.json')
    json.dump(report, open(rp, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'合并报告: {rp}（conflicts/lowconf 供仲裁遍使用）')


if __name__ == '__main__':
    main()
