# -*- coding: utf-8 -*-
"""collect_flags.py —— 汇总全部 flag，产出可直接分发给查证子代理的 flags.json。

改写遍、仲裁遍、乃至"与原文无差异但提了疑点"的段，flag 分散在十几份产物里
（pass1_part*.json / pass2_part*.json / arb.json / merged.json），历史上每次查证遍
之前都要现写 dumpflags/collect_flags 脚本人工拼。本脚本一条命令收齐：

汇总来源（按可信度排序，同 id 的多条 flag 去重后合并）:
  1. `<merged.json>` 里的 `flag`（merge_passes.py 已做过两遍去重合并，**主来源**）；
  2. 同目录 `pass*_part*.json` 里的 `flag`（保留分遍原始措辞，便于核对是否两边都提了）；
  3. `arb.json` 里的 `flag`（仲裁遍新增的疑点）。
同时标出"**无差异段**"（newText == 原文 但有 flag）——这批是纯疑点段，查证遍的重点。

分类（粗分，按关键词优先级，给查证子代理分批用）: 法规 / 标准 / 数字 / 地名 / 其他

分批：按 paraIndex 排序后每 `--batch-size` 条一批（默认 30，历史经验 ≳30 条分批并行）。

用法:
  python collect_flags.py <merged.json> [--paragraphs paragraphs.json] [-o flags.json]
      [--pass-glob "pass*_part*.json"] [--arb arb.json] [--batch-size 30] [--no-passes]
      [--respect-merged]

  例: python collect_flags.py merged.json -o flags.json
      python collect_flags.py final.json --paragraphs paragraphs.json --batch-size 20
      python collect_flags.py final.json --respect-merged     # 收口后跑：不再捡回已清空的批注

选项:
  --paragraphs F     paragraphs.json（merged 里缺某 id 时补原文）
  --pass-glob G      改写产物通配，默认 "pass*_part*.json"（相对 merged 所在目录；
                     空串或 --no-passes 表示不扫改写遍）
  --arb F            仲裁产物，默认同目录 arb.json（存在才读）
  --batch-size N     每批条数，默认 30
  -o, --out F        输出，默认 <merged 同目录>/flags.json
  --json-only        只打印一行机器可读摘要
  --respect-merged   以 merged/final 的 flag 为准：某段在 merged 里已无 flag（=finalize 清空了
                     批注）时，不再从改写遍/仲裁遍把旧 flag 捡回来。finalize 之后跑查证遍时用它
                     （实测：final.json 默认汇总 71 条，加它后 58 条，与成品批注数一致）。

输出结构:
  { "meta": {...计数...},
    "items": [{"id","paraIndex","original","newText","changed","unchanged_flag",
               "flag","flag_fragments":[{"text","source"}],"sources":[...],"type","batch"}],
    "batches": [{"batch","count","ids","by_type"}] }

退出码: 0 正常（含 0 条 flag）；1 输入文件缺失/读不出。
"""
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BATCH_DEFAULT = 30

# 分类规则：按顺序命中即停（法规/标准优先——它们有权威对照表，最值得先查证）
TYPE_RULES = [
    ('标准', re.compile(
        r'(GB|SL|DL|NB|CJ|JG|JT|QX|DB|HJ|YD|CECS|T/CECS)[\s/]*T?[\s]*\d'
        r'|标准|规范|规程|导则|条文|强制性|现行版|版本号')),
    ('法规', re.compile(
        r'《[^》]{2,40}》|法规|条例|办法|规定|通知|意见|批复|函|国务院|水利部|省政府'
        r'|年[^。；]{0,8}号|文号')),
    ('数字', re.compile(
        r'数字|数值|数据|年份|年度|流量|水位|面积|高程|库容|重现期|频率|量纲|单位'
        r'|\d{4}\s*年|\d+(?:\.\d+)?')),
    ('地名', re.compile(
        r'地名|专名|圩|村|镇|乡|县|市|区|河|湖|库|洲|站名|行政区')),
]


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def classify(flag_text):
    for name, rx in TYPE_RULES:
        if rx.search(flag_text or ''):
            return name
    return '其他'


def dedupe_join(texts):
    """去重（完全相同的只留一条；互相包含的保留信息量大的），用；连接。"""
    out = []
    for t in texts:
        t = (t or '').strip()
        if not t:
            continue
        if t in out:
            continue
        replaced = False
        for i, kept in enumerate(out):
            if t in kept:
                replaced = True
                break
            if kept in t:
                out[i] = t
                replaced = True
                break
        if not replaced:
            out.append(t)
    return '；'.join(out)


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('merged')
    ap.add_argument('--paragraphs', default='')
    ap.add_argument('--pass-glob', default='pass*_part*.json')
    ap.add_argument('--no-passes', action='store_true')
    ap.add_argument('--arb', default='')
    ap.add_argument('--batch-size', type=int, default=BATCH_DEFAULT)
    ap.add_argument('-o', '--out', default='')
    ap.add_argument('--json-only', action='store_true')
    ap.add_argument('--respect-merged', action='store_true',
                    help='merged/final 里该段已无 flag（查证后清了批注）时，不再从改写遍/仲裁遍捡回来')
    args = ap.parse_args()

    mp = Path(args.merged).resolve()
    if not mp.exists():
        print(f'✗ merged 不存在: {mp}')
        sys.exit(1)
    base = mp.parent
    try:
        merged_list = load_json(mp)
    except Exception as e:
        print(f'✗ merged 读不出: {e}')
        sys.exit(1)
    if not isinstance(merged_list, list):
        print(f'✗ merged 应为数组（rewrites 格式），实际是 {type(merged_list).__name__}')
        sys.exit(1)

    paras = {}
    if args.paragraphs:
        pp = Path(args.paragraphs).resolve()
        if pp.exists():
            paras = {p['id']: p for p in load_json(pp)}
        else:
            print(f'⚠ --paragraphs 不存在（{pp}），只在 merged 缺 id 时影响原文字段')

    # id -> 汇总记录
    rec = {}
    n_merged_flags = 0
    merged_has_flag = {}
    for m in merged_list:
        pid = m.get('id')
        orig = m.get('original')
        if orig is None and pid in paras:
            orig = paras[pid]['text']
        new = m.get('newText')
        rec[pid] = {
            'id': pid,
            'paraIndex': m.get('paraIndex', paras.get(pid, {}).get('paraIndex')),
            'original': orig if orig is not None else '',
            'newText': new if isinstance(new, str) else (orig or ''),
            'fragments': [],
        }
        fl = str(m.get('flag') or '').strip()
        merged_has_flag[pid] = bool(fl)
        if fl:
            n_merged_flags += 1
            rec[pid]['fragments'].append({'text': fl, 'source': 'merged'})

    n_pass_files = 0
    if not args.no_passes and args.pass_glob:
        for f in sorted(base.glob(args.pass_glob)):
            try:
                data = load_json(f)
            except Exception as e:
                print(f'⚠ 跳过读不出的改写产物 {f.name}: {e}')
                continue
            if isinstance(data, dict):
                data = data.get('changes') or []
            if not isinstance(data, list):
                continue
            n_pass_files += 1
            for r in data:
                pid = r.get('id')
                fl = str(r.get('flag') or '').strip()
                if not fl:
                    continue
                if pid not in rec:
                    orig = paras.get(pid, {}).get('text', '')
                    rec[pid] = {'id': pid,
                                'paraIndex': (paras.get(pid) or {}).get('paraIndex'),
                                'original': orig, 'newText': orig, 'fragments': []}
                rec[pid]['fragments'].append({'text': fl, 'source': f.stem})

    n_arb = 0
    arb_path = Path(args.arb).resolve() if args.arb else base / 'arb.json'
    if arb_path.exists():
        try:
            arb = load_json(arb_path)
            if isinstance(arb, dict):
                arb = arb.get('changes') or []
            for r in arb or []:
                pid = r.get('id')
                fl = str(r.get('flag') or '').strip()
                if not fl:
                    continue
                n_arb += 1
                if pid not in rec:
                    orig = paras.get(pid, {}).get('text', '')
                    rec[pid] = {'id': pid,
                                'paraIndex': (paras.get(pid) or {}).get('paraIndex'),
                                'original': orig, 'newText': orig, 'fragments': []}
                rec[pid]['fragments'].append({'text': fl, 'source': 'arb'})
        except Exception as e:
            print(f'⚠ 仲裁产物读不出（{arb_path.name}）: {e}')

    items = []
    n_cleared = 0
    for pid, r in rec.items():
        if not r['fragments']:
            continue
        if args.respect_merged and pid in merged_has_flag and not merged_has_flag[pid]:
            # merged（通常是收口后的 final.json）已明确清空该段 flag ⇒ 视为查证后误报，不再下发
            n_cleared += 1
            continue
        texts = []
        sources = []
        for frag in r['fragments']:
            if frag['text'] not in texts:
                texts.append(frag['text'])
            if frag['source'] not in sources:
                sources.append(frag['source'])
        combined = dedupe_join(texts)
        changed = (r['newText'] != r['original'])
        items.append({
            'id': pid,
            'paraIndex': r['paraIndex'],
            'original': r['original'],
            'newText': r['newText'],
            'changed': changed,
            'unchanged_flag': not changed,
            'flag': combined,
            'flag_fragments': r['fragments'],
            'sources': sources,
            'type': classify(combined),
        })

    items.sort(key=lambda x: (x['paraIndex'] if x['paraIndex'] is not None else 10 ** 9,
                              x['id']))
    bs = max(1, args.batch_size)
    batches = []
    for i in range(0, len(items), bs):
        chunk = items[i:i + bs]
        b = len(batches) + 1
        for it in chunk:
            it['batch'] = b
        batches.append({
            'batch': b,
            'count': len(chunk),
            'ids': [it['id'] for it in chunk],
            'by_type': dict(Counter(it['type'] for it in chunk)),
        })

    by_type = Counter(it['type'] for it in items)
    by_source = Counter(s for it in items for s in it['sources'])
    n_unchanged = sum(1 for it in items if it['unchanged_flag'])
    summary = {
        'tool': 'collect_flags.py',
        'merged': str(mp),
        'scan_pass_files': n_pass_files,
        'arb_flags': n_arb,
        'total_flags': len(items),
        'cleared_by_merged': n_cleared,
        'unchanged_flags': n_unchanged,
        'by_type': dict(by_type.most_common()),
        'by_source': dict(by_source.most_common()),
        'batches': len(batches),
        'batch_size': bs,
    }
    out = Path(args.out) if args.out else base / 'flags.json'
    payload = {'meta': summary, 'items': items, 'batches': batches}
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding='utf-8')

    if args.json_only:
        print(json.dumps(summary, ensure_ascii=False))
        return 0

    print(f'来源: merged flag {n_merged_flags} 条 + 改写遍 {n_pass_files} 份产物'
          f' + 仲裁 {n_arb} 条')
    print(f'汇总 flag: {len(items)} 段（其中无差异段 {n_unchanged} 段）')
    print('按类型: ' + ', '.join(f'{k} {v}' for k, v in by_type.most_common()))
    if n_cleared:
        print(f'（--respect-merged：{n_cleared} 条 flag 已被 merged 清空，不再下发）')
    print('按来源: ' + ', '.join(f'{k} {v}' for k, v in by_source.most_common()))
    print(f'批次建议: {len(batches)} 批（每批 ≤{bs} 条）'
          + ('' if not batches else '，各批条数 ' + str([b['count'] for b in batches])))
    if items:
        print('\n前 8 条:')
        for it in items[:8]:
            print(f'  id{it["id"]:<5} [{it["type"]}] batch{it.get("batch")} '
                  f'{"(无差异)" if it["unchanged_flag"] else "(已改动)"} '
                  f'{it["flag"][:80]}')
    print(f'\n→ {out}')
    return 0


if __name__ == '__main__':
    sys.exit(main() or 0)
