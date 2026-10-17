# -*- coding: utf-8 -*-
"""build_arb.py —— 建仲裁输入（把 merge-report 的冲突/低置信段切成 N 个仲裁包）。

之前 13 个运行目录里各自手写一份 build_arb.py（泸水第三轮还是从上一轮直接拷的），
接口、分片口径、上下文附加各不相同。本脚本固化成一个入口：

  1. 取 `<merged.merge-report.json>` 里的 `conflict`（多候选）与 `lowconf`（单候选）段；
  2. **按 id 排序**后按累计 JSON 体积均衡切成 N 包（默认 6，与改写遍分片数对齐）；
  3. 每包内每条附：段 id / paraIndex / kind / 原文 / 各候选 newText /
     **前后各 1 段的文本摘要**（模型判"本段是否自相矛盾"时需要上下文）/
     改写遍的 flag / lint 命中（都可选，有就带上）；
  4. 输出 `arb_input_part1..N.json`（JSON 数组，直接给仲裁子代理读）。

`--combine`：把仲裁子代理产出的 `arb_part*.json` 合并成 `arb.json`
（give merge_passes.py --arb 用），并做基本校验：id 唯一、newText 非空。

用法:
  # 造仲裁输入
  python build_arb.py merged.merge-report.json --paragraphs paragraphs.json \
      [--parts 6] [--outdir .] [--lint paragraphs-lint.json]
  # 合并仲裁产物
  python build_arb.py --combine [merged.merge-report.json] [--outdir .] [-o arb.json]
  例:
    python build_arb.py merged.merge-report.json --paragraphs paragraphs.json --parts 6
    python build_arb.py --combine -o arb.json

选项:
  --parts N        仲裁包数，默认 6；冲突段不足 N 条时自动降为条数
  --outdir DIR     产物目录，默认 merge-report 所在目录
  --template T     build 模式输出文件名模板，默认 "arb_input_part{}.json"
  --combine        合并模式（不造输入，只合 arb_part*.json）
  --arb-template T combine 模式输入模板，默认 "arb_part{}.json"
  --count N        combine 模式显式给包数（默认按 arb_part*.json 自动发现编号）
  -o, --out        combine 模式输出路径，默认 <outdir>/arb.json
  --no-flags       不自动带上改写遍 flag
  --lint FILE      lint 产物（可选，带上更便于仲裁判断）

退出码: 0 正常；1 输入缺失/无冲突段可造包之外的硬错误，或 combine 校验失败。
"""
import argparse
import glob
import json
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

CTX_CHARS = 60          # 上下文摘要长度（前后各 1 段）


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def save_json(path, obj, indent=1):
    Path(path).write_text(json.dumps(obj, ensure_ascii=False, indent=indent),
                          encoding='utf-8')


def brief(text, n=CTX_CHARS):
    text = (text or '').strip()
    return text if len(text) <= n else text[:n] + '…'


# ---------------------------------------------------------------- 造输入
def collect_items(report, paras_list, lint_by_id, flags_by_id):
    """把 report 的 conflict/lowconf 变成仲裁条目（按 id 排序）。"""
    para_by_id = {}
    pos_by_id = {}
    for i, p in enumerate(paras_list):
        para_by_id[p['id']] = p
        pos_by_id[p['id']] = i

    def ctx(pid):
        pos = pos_by_id.get(pid)
        out = {'prev': None, 'next': None}
        if pos is None:
            return out
        if pos > 0:
            q = paras_list[pos - 1]
            out['prev'] = {'id': q['id'], 'text': brief(q['text'])}
        if pos + 1 < len(paras_list):
            q = paras_list[pos + 1]
            out['next'] = {'id': q['id'], 'text': brief(q['text'])}
        return out

    items, unknown = [], []
    for kind in ('conflict', 'lowconf'):
        for e in report.get(kind) or []:
            pid = e.get('id')
            p = para_by_id.get(pid)
            if p is None:
                unknown.append(pid)
                continue
            cands = e.get('candidates') if kind == 'conflict' else [e.get('candidate')]
            cands = [str(c) for c in (cands or [])]
            item = {
                'id': pid,
                'paraIndex': p.get('paraIndex'),
                'kind': kind,
                'original': p['text'],
                'candidates': cands,
                'context': ctx(pid),
                'flags': flags_by_id.get(pid, []),
                'lint': [i for i in (lint_by_id.get(pid) or {}).get('issues', [])],
            }
            items.append(item)
    items.sort(key=lambda x: x['id'])
    return items, unknown


def split_by_size(items, n):
    """按累计 JSON 体积均衡分片；每片非空（n > 条数时降为条数）。"""
    if not items:
        return []
    n = max(1, min(n, len(items)))
    total = sum(len(json.dumps(i, ensure_ascii=False)) for i in items)
    target = total / n
    parts, cur, cur_sz = [], [], 0
    for it in items:
        cur.append(it)
        cur_sz += len(json.dumps(it, ensure_ascii=False))
        if cur_sz >= target and len(parts) < n - 1:
            parts.append(cur)
            cur, cur_sz = [], 0
    if cur:
        parts.append(cur)
    return parts


def cmd_build(args):
    if not args.report:
        print('✗ build 模式需要 <merged.merge-report.json> 作为第一个参数')
        return 1
    report_path = Path(args.report).resolve()
    if not report_path.exists():
        print(f'✗ merge-report 不存在: {report_path}')
        return 1
    if not args.paragraphs:
        print('✗ build 模式必须给 --paragraphs paragraphs.json（原文与上下文都从它取）')
        return 1
    pj = Path(args.paragraphs).resolve()
    if not pj.exists():
        print(f'✗ paragraphs 不存在: {pj}')
        return 1
    outdir = Path(args.outdir).resolve() if args.outdir else report_path.parent
    outdir.mkdir(parents=True, exist_ok=True)

    report = load_json(report_path)
    paras_list = load_json(pj)
    lint_by_id = {}
    if args.lint:
        lp = Path(args.lint).resolve()
        if lp.exists():
            lint_by_id = {x['id']: x for x in load_json(lp)}
        else:
            print(f'⚠ --lint 不存在（{lp}），按无 lint 处理')

    flags_by_id = {}
    if not args.no_flags:
        n_pass = 0
        for f in sorted(outdir.glob('pass*_part*.json')):
            try:
                data = load_json(f)
            except Exception as e:
                print(f'⚠ 跳过读不出的改写产物 {f.name}: {e}')
                continue
            if isinstance(data, dict):
                data = data.get('changes') or []
            for r in data:
                fl = str(r.get('flag') or '').strip()
                if fl:
                    flags_by_id.setdefault(r.get('id'), [])
                    if fl not in flags_by_id[r.get('id')]:
                        flags_by_id[r.get('id')].append(fl)
            n_pass += 1
        if n_pass:
            print(f'· 自动带上 {n_pass} 份改写产物的 flag（{len(flags_by_id)} 段有旗标）')

    items, unknown = collect_items(report, paras_list, lint_by_id, flags_by_id)
    if unknown:
        print(f'⚠ merge-report 里 {len(unknown)} 个 id 不在 paragraphs.json 中: {unknown[:8]}')
    n_conflict = len(report.get('conflict') or [])
    n_lowconf = len(report.get('lowconf') or [])
    print(f'冲突段 {n_conflict} + 低置信段 {n_lowconf} = 仲裁条目 {len(items)}')
    if not items:
        print('· 无冲突/低置信段，本轮不需要仲裁遍（未写任何 arb_input 文件）')
        return 0

    parts = split_by_size(items, args.parts)
    total_items = 0
    for i, chunk in enumerate(parts, 1):
        path = outdir / args.template.format(i)
        save_json(path, chunk)
        total_items += len(chunk)
        kinds = {}
        for it in chunk:
            kinds[it['kind']] = kinds.get(it['kind'], 0) + 1
        chars = len(json.dumps(chunk, ensure_ascii=False))
        print(f'  {path.name}: {len(chunk)} 条 / {chars} 字符 / '
              f'id {chunk[0]["id"]}..{chunk[-1]["id"]} / {kinds}')
    print(f'→ 共 {len(parts)} 包 {total_items} 条，目录 {outdir}')
    print('下一步：把 arb_input_part*.json 分发给仲裁子代理（模板 '
          'scripts/templates/PROMPT_arb.md），产出后用 check_arb.py 自检')
    return 0


# ---------------------------------------------------------------- 合并
def cmd_combine(args):
    outdir = Path(args.outdir).resolve() if args.outdir else Path.cwd()
    if args.report:
        rp = Path(args.report).resolve()
        if rp.exists():
            outdir = Path(args.outdir).resolve() if args.outdir else rp.parent
    # 发现分片编号
    if args.count:
        paths = [outdir / args.arb_template.format(i) for i in range(1, args.count + 1)]
    else:
        found = {}
        pat = re.compile('^' + re.escape(args.arb_template).replace(r'\{\}', r'(\d+)') + '$')
        for f in outdir.glob('*.json'):
            m = pat.match(f.name)
            if m:
                found[int(m.group(1))] = f
        paths = [found[k] for k in sorted(found)]
        if not paths:
            # 兜底：按 arb_part*.json 通配（模板被改过时仍能用）
            for f in outdir.glob('arb_part*.json'):
                m = re.search(r'(\d+)\.json$', f.name)
                if m:
                    found[int(m.group(1))] = f
            paths = [found[k] for k in sorted(found)]

    if not paths:
        print(f'✗ 在 {outdir} 下没找到任何仲裁产物（{args.arb_template}）')
        return 1
    missing = [p.name for p in paths if not p.exists()]
    if missing:
        print(f'✗ 仲裁分包缺失: {missing}')
        return 1

    # 若给了 report，用它校验 id 归属
    allowed_ids, cand_map = set(), {}
    if args.report:
        rp = Path(args.report).resolve()
        if rp.exists():
            rep = load_json(rp)
            for e in rep.get('conflict') or []:
                allowed_ids.add(e.get('id'))
                cand_map[e.get('id')] = [str(c) for c in (e.get('candidates') or [])]
            for e in rep.get('lowconf') or []:
                allowed_ids.add(e.get('id'))
                cand_map[e.get('id')] = [str(e.get('candidate'))]
        else:
            print(f'⚠ --combine 的 report 不存在（{rp}），跳过 id 归属校验')

    errors, seen, merged = [], {}, []
    for path in paths:
        try:
            data = load_json(path)
        except Exception as e:
            errors.append(f'{path.name}: JSON 解析失败 {e}')
            continue
        if isinstance(data, dict):
            data = data.get('changes') or []
        if not isinstance(data, list):
            errors.append(f'{path.name}: 顶层不是数组')
            continue
        if not data:
            errors.append(f'{path.name}: 空数组（仲裁产物不应为空）')
            continue
        for r in data:
            pid = r.get('id')
            if pid is None:
                errors.append(f'{path.name}: 有条目缺 id')
                continue
            if pid in seen:
                errors.append(f'id{pid} 重复（{seen[pid].name} 与 {path.name}）')
                continue
            seen[pid] = path
            nt = r.get('newText')
            if not isinstance(nt, str) or not nt.strip():
                errors.append(f'id{pid}: newText 为空/非字符串')
                continue
            if re.search(r'[\r\n]', nt):
                errors.append(f'id{pid}: newText 含换行')
            if allowed_ids and pid not in allowed_ids:
                errors.append(f'id{pid}: 不在 merge-report 的冲突/低置信段中')
            entry = {'id': pid, 'newText': nt.strip()}
            if str(r.get('flag') or '').strip():
                entry['flag'] = str(r['flag']).strip()
            merged.append(entry)

    if errors:
        print(f'✗ 校验失败 {len(errors)} 处：')
        for e in errors[:30]:
            print('  - ' + e)
        return 1

    merged.sort(key=lambda x: x['id'])
    # 统计：采纳候选 vs 保持原文（有 report 时）
    n_keep = n_cand = 0
    if cand_map:
        for e in merged:
            cands = cand_map.get(e['id']) or []
            if cands and e['newText'] in cands:
                n_cand += 1
            else:
                n_keep += 1
    out = Path(args.out).resolve() if args.out else outdir / 'arb.json'
    save_json(out, merged)
    print(f'合并 {len(paths)} 包 → {len(merged)} 条 → {out}')
    if cand_map:
        print(f'  采纳候选 {n_cand} 条；保持原文/非候选 {n_keep} 条')
    print('下一步：merge_passes.py ... --arb arb.json')
    return 0


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('report', nargs='?', default='',
                    help='merged.merge-report.json（build 模式必填；combine 模式可选，用于校验）')
    ap.add_argument('--paragraphs', default='', help='paragraphs.json（原文与上下文来源）')
    ap.add_argument('--parts', type=int, default=6, help='仲裁包数，默认 6')
    ap.add_argument('--outdir', default='', help='产物目录，默认 merge-report 所在目录')
    ap.add_argument('--template', default='arb_input_part{}.json',
                    help='build 模式输出文件名模板')
    ap.add_argument('--combine', action='store_true', help='合并模式：合 arb_part*.json → arb.json')
    ap.add_argument('--arb-template', default='arb_part{}.json', help='combine 模式输入模板')
    ap.add_argument('--count', type=int, default=0, help='combine 模式显式包数')
    ap.add_argument('-o', '--out', default='', help='combine 模式输出路径（默认 arb.json）')
    ap.add_argument('--no-flags', action='store_true', help='不自动带改写遍 flag')
    ap.add_argument('--lint', default='', help='lint 产物（可选）')
    args = ap.parse_args()
    return cmd_combine(args) if args.combine else cmd_build(args)


if __name__ == '__main__':
    sys.exit(main() or 0)
