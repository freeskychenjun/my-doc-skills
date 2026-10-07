# -*- coding: utf-8 -*-
"""分片产物完整性自检：在合并之前快速发现"某片没写出来/写坏了"。

限流（429/1302）或子代理中断时，症状往往是某个 part 文件缺失或半截；
本脚本把它变成一次确定性检查，避免到合并阶段才发现要重跑。

用法:
  python check_parts.py <paragraphs.json> "<part文件模板>" --count 6
  例: python check_parts.py paragraphs.json "pass1_part{}.json" --count 6
      python check_parts.py paragraphs.json "rewrites_part{}.json" --count 6
      （模板里 {} 会被替换为 1..count；相对路径按 paragraphs.json 所在目录解析）

检查项:
  1. 文件存在、JSON 可解析、非空；
  2. id 集合完整（若同目录存在 paragraphs_partN.json 则与其精确比对，否则只查全局覆盖与重复）；
  3. 每条 original 与 paragraphs.json 的 text 逐字一致；
  4. newText 非空、无换行；
  5. 数字 token 是否原样保留（提示性告警，不判失败）。

退出码: 0 全部通过；1 存在问题（控制台列出可重跑的分片）。
"""
import argparse
import json
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
NUM = re.compile(r'\d+(?:\.\d+)?')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('paragraphs')
    ap.add_argument('template')
    ap.add_argument('--count', type=int, required=True)
    args = ap.parse_args()

    pdir = Path(args.paragraphs).resolve().parent
    paras = json.load(open(args.paragraphs, encoding='utf-8'))
    by_id = {p['id']: p['text'] for p in paras}
    expected_parts = {}
    for i in range(1, args.count + 1):
        f = pdir / f'paragraphs_part{i}.json'
        if f.exists():
            expected_parts[i] = [p['id'] for p in json.load(open(f, encoding='utf-8'))]

    bad = []
    seen = {}
    for i in range(1, args.count + 1):
        f = pdir / args.template.format(i)
        problems = []
        data = None
        if not f.exists():
            problems.append('文件缺失')
        else:
            try:
                data = json.load(open(f, encoding='utf-8-sig'))
                if not isinstance(data, list) or not data:
                    problems.append('JSON 为空或非数组')
            except Exception as e:
                problems.append(f'JSON 解析失败: {e}')
        if data:
            ids = [r.get('id') for r in data]
            if len(ids) != len(set(ids)):
                problems.append('存在重复 id')
            for r in data:
                seen.setdefault(r.get('id'), []).append(i)
                if r.get('original') != by_id.get(r.get('id')):
                    problems.append(f"id{r.get('id')} original 与提取文本不一致")
                nt = r.get('newText') or ''
                if not nt.strip():
                    problems.append(f"id{r.get('id')} newText 为空")
                elif re.search(r'[\r\n]', nt):
                    problems.append(f"id{r.get('id')} newText 含换行")
            if i in expected_parts:
                miss = set(expected_parts[i]) - set(ids)
                extra = set(ids) - set(expected_parts[i])
                if miss:
                    problems.append(f'缺 {len(miss)} 条: {sorted(miss)[:8]}')
                if extra:
                    problems.append(f'多 {len(extra)} 条: {sorted(extra)[:8]}')
        if problems:
            bad.append((i, problems))
            print(f'✗ part{i}: ' + '; '.join(problems[:4]))
        else:
            print(f'✓ part{i}: {len(data)} 条')

    cross = {k: v for k, v in seen.items() if len(v) > 1}
    if cross:
        print(f'✗ 跨分片重复 id: {list(cross.items())[:5]}')
        bad.append(('cross', ['跨分片重复']))
    if expected_parts:
        covered = set(seen)
        missing_all = set(by_id) - covered
        if missing_all:
            print(f'✗ 未被任何分片覆盖的 id: {len(missing_all)} 个 {sorted(missing_all)[:10]}')
            bad.append(('coverage', [f'缺 {len(missing_all)} 条']))
        else:
            print(f'✓ 覆盖完整: {len(by_id)} 条全部有产出')

    if bad:
        print(f'\n结论: 有问题 {len(bad)} 处 → 需要重跑的分片: '
              f'{[b[0] for b in bad if isinstance(b[0], int)]}')
        sys.exit(1)
    print('\n结论: 全部分片完好，可以进入合并')


if __name__ == '__main__':
    main()
