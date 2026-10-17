# -*- coding: utf-8 -*-
"""check_arb.py —— 仲裁产物自检（逐字候选断言）。

历史上 8 个运行目录各抄一份 selfcheck_arb.py，核心断言只有一句：
`newText in candidates` —— 防的是**仲裁 agent 自己改写候选**（明令禁止：父代理要的是
"从既有候选里挑一条"，自写会绕过两遍投票的质量控制）。本脚本把它固化，并补上
id 归属、跨包重复、覆盖率三类检查。

检查项（逐包）:
  1. 文件存在、JSON 可解析为数组、非空；
  2. 每条 id 必须属于**本包**的 id 集合（若同目录有 arb_input_part{N}.json 则以它为准，
     否则退化为"必须属于 merge-report 的 conflict+lowconf 全集"）；
  3. 本包内 id 不重复；跨包 id 不重复（硬失败）；
  4. `newText` 必须**逐字等于**该段的某条候选（candidates）**或原文（original）**。
     —— 允许等于原文是刻意的：仲裁的合法结论之一是"保持原文不采纳候选"
        （泸水第三轮实测 124/195 条就是保持原文；严格只认 candidates 会把这批
        合法裁决全判失败）。要强制"必须落在候选里"加 `--require-candidate`。
        **既不在候选、也不是原文 ⇒ 判"仲裁自改"，硬失败。**
  5. newText 非空串、不含换行；flag（若有）为字符串；
  6. 覆盖率（**不判失败**，只报告）：哪些 conflict/lowconf 段没出现在任何包里。

用法:
  python check_arb.py <merged.merge-report.json> "arb_part{}.json" --count 6
  例: python check_arb.py merged.merge-report.json "arb_part{}.json" --count 6
      python check_arb.py merged.merge-report.json "arb_part{}.json" --count 6 \
          --input-template "arb_input_part{}.json" --json arb-check.json

选项:
  --count N            包数（包文件编号 1..N，缺一个即硬失败）
  --input-template T   仲裁输入模板（默认 "arb_input_part{}.json"；存在则用于精确的
                       包内 id 集合、原文与候选；不存在则回落到 merge-report）
  --require-candidate  严格模式：newText 必须落在 candidates 里（"保持原文"也判失败）
  --json FILE          写机器可读摘要
  -v                   打印每个被判定"保持原文"的 id

退出码: 0 全部通过；1 有硬失败（控制台列出问题与涉及的分包）。
"""
import argparse
import json
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def build_expected(report, part_dir, count, input_template):
    """返回 ({part: set(ids)}, {id: {candidates, original, kind}})"""
    expected_ids, meta = {}, {}
    for kind in ('conflict', 'lowconf'):
        for e in report.get(kind) or []:
            pid = e.get('id')
            cands = e.get('candidates') if kind == 'conflict' else [e.get('candidate')]
            meta.setdefault(pid, {'kind': kind, 'candidates': [], 'original': None})
            meta[pid]['candidates'] = [str(c) for c in (cands or [])]
    # 优先用仲裁输入包（它自带 original，且包划分就是父代理当时的划分）
    found = 0
    for i in range(1, count + 1):
        f = part_dir / input_template.format(i)
        if not f.exists():
            continue
        try:
            data = load_json(f)
        except Exception:
            continue
        if not isinstance(data, list):
            continue
        found += 1
        expected_ids[i] = {it['id'] for it in data}
        for it in data:
            pid = it['id']
            meta.setdefault(pid, {'kind': it.get('kind'), 'candidates': [], 'original': None})
            meta[pid]['kind'] = it.get('kind') or meta[pid].get('kind')
            meta[pid]['candidates'] = [str(c) for c in (it.get('candidates') or [])]
            meta[pid]['original'] = it.get('original')
    if found == 0:
        # 回落：只认"属于全集"，不校验包划分
        for i in range(1, count + 1):
            expected_ids[i] = set(meta)
    return expected_ids, meta, found


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('report', help='merged.merge-report.json')
    ap.add_argument('template', help='仲裁产物文件名模板，如 "arb_part{}.json"')
    ap.add_argument('--count', type=int, required=True, help='包数')
    ap.add_argument('--input-template', default='arb_input_part{}.json',
                    help='仲裁输入模板（默认 arb_input_part{}.json）')
    ap.add_argument('--require-candidate', action='store_true',
                    help='严格模式：newText 必须落在 candidates 里（保持原文也判失败）')
    ap.add_argument('--json', default='', help='写机器可读摘要')
    ap.add_argument('-v', '--verbose', action='store_true')
    args = ap.parse_args()

    rp = Path(args.report).resolve()
    if not rp.exists():
        print(f'✗ merge-report 不存在: {rp}')
        sys.exit(1)
    report = load_json(rp)
    part_dir = rp.parent
    expected_ids, meta, n_in = build_expected(report, part_dir, args.count,
                                              args.input_template)
    total_expected = len(meta)
    print(f'合并报告: {rp.name}')
    print(f'冲突/低置信段共 {total_expected} 条；仲裁输入包 {n_in}/{args.count} 个'
          + ('' if n_in else '（未找到输入包，包内 id 集合校验退化为全集校验）'))
    if total_expected == 0:
        print('⚠ merge-report 里没有 conflict/lowconf 段，且同目录未找到仲裁输入包：'
              '没有候选可比对，"逐字候选断言"退化为无校验（报告可能已在仲裁后被重算过）')

    hard, soft = [], []
    adopted = kept_orig = 0
    seen = {}
    covered = set()
    part_counts = {}
    for part in range(1, args.count + 1):
        path = part_dir / args.template.format(part)
        if not path.exists():
            hard.append(f'part{part}: 文件缺失（{path.name}）')
            continue
        try:
            data = load_json(path)
        except Exception as e:
            hard.append(f'part{part}: JSON 解析失败 {e}')
            continue
        if not isinstance(data, list):
            hard.append(f'part{part}: 顶层不是数组')
            continue
        exp = expected_ids.get(part, set())
        if exp and not data:
            hard.append(f'part{part}: 空数组（本包应有 {len(exp)} 条待裁定）')
            continue
        local = {}
        local_hard = []
        n_adopt = n_keep = 0

        def bad(msg):
            hard.append(msg)
            local_hard.append(msg)

        for r in data:
            pid = r.get('id')
            if pid is None:
                bad(f'part{part}: 有条目缺 id（{str(r)[:60]}）')
                continue
            if pid in local:
                bad(f'part{part}: 包内 id{pid} 重复')
                continue
            local[pid] = True
            if pid in seen:
                bad(f'part{part}: 跨包 id{pid} 重复（part{seen[pid]} 与 part{part}）')
                continue
            seen[pid] = part
            covered.add(pid)
            if exp and pid not in exp:
                bad(f'part{part}: id{pid} 不属于本包（本包 id 范围 '
                    f'{min(exp) if exp else "-"}..{max(exp) if exp else "-"}）')
                continue
            if total_expected and pid not in meta:
                bad(f'part{part}: id{pid} 不在 merge-report 的冲突/低置信段里')
                continue
            nt = r.get('newText')
            if not isinstance(nt, str) or not nt.strip():
                bad(f'part{part}: id{pid} newText 为空/非字符串')
                continue
            if re.search(r'[\r\n]', nt):
                bad(f'part{part}: id{pid} newText 含换行')
            info = meta.get(pid, {}) or {}
            cands = info.get('candidates') or []
            orig = info.get('original')
            if cands and nt in cands:
                n_adopt += 1
            elif orig is not None and nt == orig:
                n_keep += 1
                if args.require_candidate:
                    bad(f'part{part}: id{pid} newText 等于原文，未采纳任何候选'
                        f'（--require-candidate 严格模式下判失败）')
                elif args.verbose:
                    soft.append(f'part{part}: id{pid} 保持原文（合法裁决）')
            elif cands or orig is not None:
                bad(f'part{part}: id{pid} **仲裁自改**：newText 既非 candidates '
                    f'中任何一条、也不等于原文（禁止自己重写；候选 {len(cands)} 条）')
            flag = r.get('flag')
            if flag is not None and not isinstance(flag, str):
                bad(f'part{part}: id{pid} flag 不是字符串')
        adopted += n_adopt
        kept_orig += n_keep
        part_counts[part] = {'items': len(data), 'adopted': n_adopt, 'keep_original': n_keep}
        print(f'✓ part{part}: {len(data)} 条（采纳候选 {n_adopt} / 保持原文 {n_keep}）'
              if not local_hard
              else f'✗ part{part}: {len(data)} 条（{len(local_hard)} 处问题，见下）')

    for s in soft[:20]:
        print('  · ' + s)

    missing = []
    if total_expected:
        missing = sorted(set(meta) - covered)
    if missing:
        print(f'\n覆盖率：{len(covered)}/{total_expected} 条被仲裁；'
              f'未覆盖 {len(missing)} 条（**不判失败**，通常是仲裁 agent 漏条）:')
        print('  ' + str(missing[:40]) + (' …' if len(missing) > 40 else ''))
    elif total_expected:
        print(f'\n覆盖率：{total_expected}/{total_expected} 条全部有裁定')

    summary = {
        'ok': not hard,
        'report': str(rp),
        'count': args.count,
        'expected': total_expected,
        'arbitrated': len(covered),
        'adopted_candidate': adopted,
        'keep_original': kept_orig,
        'missing_ids': missing,
        'part_counts': part_counts,
        'hard_failures': hard,
        'warnings': soft,
    }
    if args.json:
        Path(args.json).write_text(json.dumps(summary, ensure_ascii=False, indent=1),
                                   encoding='utf-8')
        print(f'摘要 → {args.json}')

    if hard:
        print(f'\n结论: 有 {len(hard)} 处硬失败 → 仲裁产物不可用，需重跑对应分包')
        for h in hard[:30]:
            print('  ✗ ' + h)
        if len(hard) > 30:
            print(f'  …共 {len(hard)} 条')
        sys.exit(1)
    print('\n结论: 仲裁产物通过自检（逐字候选断言 / id 归属 / 无跨包重复 / 覆盖率）')
    return 0


if __name__ == '__main__':
    sys.exit(main() or 0)
