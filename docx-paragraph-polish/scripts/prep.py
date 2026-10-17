# -*- coding: utf-8 -*-
"""prep.py —— 备料一条龙（把 P0 阶段的 5 步手动操作压成 1 步）。

以前每轮 P0 都要手工串 5 条命令（extract → lint → check_refs → check_figures_tables
→ split_parts），并且 split_parts.py 每轮现抄。本脚本一次调用全部做完，并把结果汇总成
`prep-report.json` + 一份中文待办清单（父代理照着清单走，不必再看中间产物）。

顺序与产物（全部落在 --workdir，默认 docx 所在目录）：
  1. extract_body_paragraphs.py  → paragraphs.json
  2. lint_paragraphs.py          → paragraphs-lint.json
  3. check_refs.py               → paragraphs-refcheck.json（法规/标准引用核查）
  4. check_figures_tables.py     → <docx stem>-figcheck.json（图表编号/引用核查）
  5. 按累计字数分片              → paragraphs_part1..N.json + lint_part1..N.json
                                   + parts-manifest.txt
  6. 汇总                        → prep-report.json（含分片统计/各类命中数/建议并发档位）

幂等：
  · 同一个 docx（路径 + 大小 + mtime + 分片数不变）重复跑，全部步骤标记 `skip`，不重建分片；
  · 某一步的产物文件缺失/被删，只补跑那一步；
  · `--force` 强制全部重跑（换分片数时自动视为变更）。
  · 判断"是否同一份文档"用 docx 的 size+mtime（内容级 hash 对 19 MB 文档太贵）。

用法:
  python prep.py <docx> [--workdir DIR] [--parts N] [--force] [--lenient]
  例: python prep.py "src.docx"
      python prep.py "src.docx" --workdir D:\\DSH工作区\\xxx_polish --parts 6

选项:
  --parts N    分片数，默认 6（按累计字数均衡；N > 段落数 时自动降到段落数）
  --force      忽略幂等状态，全部重跑
  --lenient    lint / check_refs / check_figures_tables 失败时只告警不中止
               （extract 与分片失败始终中止——它们没有产物就没法继续）
  --skip-refs        跳过 check_refs（本机没有权威对照表时的降级开关）
  --skip-figures     跳过 check_figures_tables（文档巨大、赶时间时的降级开关）

退出码: 0 全部成功（含"已是最新、跳过"）；1 某一步失败（控制台与 prep-report.json 指明是哪一步）。
"""
import argparse
import json
import math
import os
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SCRIPTS = Path(__file__).resolve().parent
DEFAULT_PARTS = 6
REPORT_NAME = 'prep-report.json'

# 每轮都会用到的 skill 自带脚本（相对本文件所在目录）
TOOL_EXTRACT = SCRIPTS / 'extract_body_paragraphs.py'
TOOL_LINT = SCRIPTS / 'lint_paragraphs.py'
TOOL_REFS = SCRIPTS / 'check_refs.py'
TOOL_FIGS = SCRIPTS / 'check_figures_tables.py'


# ---------------------------------------------------------------- 基础工具
def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def save_json(path, obj):
    Path(path).write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding='utf-8')


def run_step(name, cmd, cwd=None):
    """跑一个子脚本，返回 (ok, detail)。失败时把最后几行输出打出来。"""
    cmd = [str(c) for c in cmd]
    print(f'[{name}] {Path(cmd[1]).name} ' + ' '.join(cmd[2:]))
    env = dict(os.environ)
    env['PYTHONIOENCODING'] = 'utf-8'
    t0 = time.time()
    try:
        p = subprocess.run([sys.executable] + cmd, capture_output=True, text=True,
                           encoding='utf-8', errors='replace', env=env,
                           cwd=str(cwd) if cwd else None)
    except Exception as e:                     # 解释器都起不来
        print(f'  ✗ 无法启动: {e}')
        return False, {'cmd': cmd, 'exit': -1, 'seconds': 0.0, 'error': str(e)}
    dt = time.time() - t0
    out_tail = (p.stdout or '').strip().splitlines()[-8:]
    err_tail = (p.stderr or '').strip().splitlines()[-12:]
    ok = p.returncode == 0
    print(('  ' if ok else '  ✗ ') + f'exit={p.returncode} {dt:.1f}s')
    for line in out_tail:
        print('    ' + line)
    if not ok:
        for line in err_tail:
            print('    ! ' + line)
    return ok, {'cmd': cmd, 'exit': p.returncode, 'seconds': round(dt, 1),
                'stdout_tail': out_tail, 'stderr_tail': err_tail}


def docx_signature(docx, parts):
    st = Path(docx).stat()
    return {'docx': str(docx), 'size': st.st_size,
            'mtime_ns': st.st_mtime_ns, 'parts': parts}


# ---------------------------------------------------------------- 分片
def split_parts(paras, lint_by_id, n):
    """按累计字数把段落均分成 n 片（与 split_parts.py 同算法，向后兼容旧产物）。

    返回 [{part, paras, chars, id_first, id_last, lint}]，entries 里带 'items'。
    算法：target = 总字数 / n；累计到 target × 已分配片数 就把下一段推进新片；
    同时保证剩余段落数足够填满剩余片（否则尾部会出现空片）。
    """
    n = max(1, min(n, len(paras)))
    total = sum(len(p['text']) for p in paras)
    target = total / n
    chunks = [[] for _ in range(n)]
    acc = 0
    idx = 0
    for k, p in enumerate(paras):
        remaining_paras = len(paras) - k
        remaining_parts = n - idx
        if idx < n - 1 and (acc >= target * (idx + 1) or remaining_paras <= remaining_parts):
            idx += 1
        chunks[idx].append(p)
        acc += len(p['text'])
    stats = []
    for i, chunk in enumerate(chunks, 1):
        lint = [lint_by_id[p['id']] for p in chunk if p['id'] in lint_by_id]
        stats.append({
            'part': i,
            'items': chunk,
            'lint': lint,
            'paras': len(chunk),
            'chars': sum(len(p['text']) for p in chunk),
            'id_first': chunk[0]['id'] if chunk else None,
            'id_last': chunk[-1]['id'] if chunk else None,
            'lint_hits': sum(len(x.get('issues') or []) for x in lint),
        })
    return stats


def write_parts(workdir, stats):
    workdir = Path(workdir)
    lines = []
    total_paras = sum(s['paras'] for s in stats)
    total_chars = sum(s['chars'] for s in stats)
    for s in stats:
        i = s['part']
        save_json(workdir / f'paragraphs_part{i}.json', s['items'])
        save_json(workdir / f'lint_part{i}.json', s['lint'])
        lines.append(f'part{i}: {s["paras"]} paras, {s["chars"]} chars, '
                     f'id {s["id_first"]}..{s["id_last"]}, lint {len(s["lint"])}')
    manifest = (f'total paras={total_paras} chars={total_chars} parts={len(stats)}\n'
                + '\n'.join(lines) + '\n')
    (workdir / 'parts-manifest.txt').write_text(manifest, encoding='utf-8')
    return manifest


# ---------------------------------------------------------------- 统计
def summarize_lint(path):
    d = load_json(path)
    if not isinstance(d, list):
        raise ValueError(f'{path}: lint 产物应为数组')
    by_rule = Counter()
    for item in d:
        for it in item.get('issues') or []:
            by_rule[it.get('rule', '?')] += 1
    return {'file': str(path), 'paragraphs_hit': len(d),
            'issues': sum(len(x.get('issues') or []) for x in d),
            'by_rule': dict(by_rule.most_common(20))}


def summarize_refs(path):
    d = load_json(path)
    if not isinstance(d, list):
        raise ValueError(f'{path}: refcheck 产物应为数组')
    by_status = Counter(h.get('status', '?') for h in d if isinstance(h, dict))
    return {'file': str(path), 'issues': len(d), 'by_status': dict(by_status)}


def summarize_figs(path):
    d = load_json(path)
    if not isinstance(d, dict) or not isinstance(d.get('issues'), list):
        raise ValueError(f'{path}: figcheck 产物应为 {{file,captions,tables,img_paras,issues}} 对象')
    by_cat = Counter(i.get('category', '?') for i in d['issues'])
    by_level = Counter(i.get('level', '?') for i in d['issues'])
    return {'file': str(path), 'issues': len(d['issues']),
            'captions': d.get('captions'), 'tables': d.get('tables'),
            'img_paras': d.get('img_paras'),
            'by_category': dict(by_cat.most_common(20)),
            'by_level': dict(by_level)}


def suggest_concurrency(parts):
    """建议并发档位（口径来自实测）。

    - 6 片全并发实测零限流（泸水第三轮、锦江）；
    - 端点抖动/429 时降到 3–5（赣江实测 3–5 稳）；
    - 片数 > 6 时按 6 并发分批。
    仲裁遍的条数远少于改写遍，一并给 6。
    """
    if parts <= 3:
        note = f'分片少（{parts} 片），全并发'
        return {'pass1': parts, 'pass2': parts, 'arb': parts, 'note': note}
    conc = 6 if parts >= 6 else parts
    batches = math.ceil(parts / conc) if conc else 1
    note = (f'每遍 {conc} 并发' + (f'，{batches} 批' if batches > 1 else '（全片并发）')
            + '；实测 6 并发零限流，若端点报 429/1302 降到 3–5')
    return {'pass1': conc, 'pass2': conc, 'arb': conc, 'note': note}


def build_todo(docx, workdir, paras_n, chars, parts_n, lint, refs, figs, conc):
    """生成中文待办清单（父代理照着走，不必再翻中间产物）。"""
    wd = str(workdir)
    todo = [
        f'保护清单：python scan_protect.py "{docx}" {wd}\\paragraphs.json '
        f'-o {wd}\\protected_ids.json --fmt-out {wd}\\format_scan.json',
        f'派发两遍改写：用 scripts/templates/PROMPT_pass1.md / PROMPT_pass2.md 填占位符，'
        f'各 {parts_n} 个子代理并行（并发 {conc["pass1"]}）',
        f'分片自检（两遍各一次）：python check_parts.py {wd}\\paragraphs.json '
        f'"pass1_part{{}}.json" --count {parts_n}（pass2 同理）',
        '合并投票：python merge_passes.py <paragraphs.json> -o merged.json --passes '
        '<pass1_part1..N.json pass2_part1..N.json>',
        '仲裁三连：build_arb.py（造 arb_input_part*.json）→ 派发仲裁 → '
        'check_arb.py（逐字候选断言）→ build_arb.py --combine（合 arb.json）',
        '收口：finalize.py --merged merged.json --src <docx>（先 check_guard.py 确认回退计数为 0）',
        '写回：apply_track_changes.py <docx> final.json polish <出件.docx> '
        '--protect protected_ids.json --numeric-guard --num-exempt <id,id>',
        '终检：postapply_check.py <出件.docx> final.json --src <docx>',
        '查证遍：collect_flags.py merged.json（产出 flags.json 供查证子代理分批）',
    ]
    head = [
        f'正文段落 {paras_n} 段 / {chars} 字 → 分 {parts_n} 片',
        f'lint 确定性命中 {lint["issues"]} 条（{lint["paragraphs_hit"]} 段）',
        f'法规/标准引用问题 {refs["issues"]} 条',
        f'图表编号/引用问题 {figs["issues"]} 条',
        f'建议并发档位：pass1/pass2/仲裁 各 {conc["pass1"]}（{conc["note"]}）',
    ]
    return head, todo


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('docx')
    ap.add_argument('--workdir', default='', help='工作目录（默认 docx 所在目录）')
    ap.add_argument('--parts', type=int, default=DEFAULT_PARTS, help='分片数，默认 6')
    ap.add_argument('--force', action='store_true', help='忽略幂等状态，全部重跑')
    ap.add_argument('--lenient', action='store_true',
                    help='lint/引用/图注核查失败只告警（extract 与分片仍硬失败）')
    ap.add_argument('--skip-refs', action='store_true')
    ap.add_argument('--skip-figures', action='store_true')
    args = ap.parse_args()

    docx = Path(args.docx).resolve()
    if not docx.exists():
        print(f'✗ docx 不存在: {docx}')
        sys.exit(1)
    if args.parts < 1:
        print(f'✗ --parts 必须 ≥1（收到 {args.parts}）')
        sys.exit(1)
    workdir = Path(args.workdir).resolve() if args.workdir else docx.parent
    workdir.mkdir(parents=True, exist_ok=True)

    for tool in (TOOL_EXTRACT, TOOL_LINT):
        if not tool.exists():
            print(f'✗ 找不到配套脚本: {tool}（本脚本必须与本 skill 的 scripts/ 目录放在一起）')
            sys.exit(1)

    p_paras = workdir / 'paragraphs.json'
    p_lint = workdir / 'paragraphs-lint.json'
    p_refs = workdir / 'paragraphs-refcheck.json'
    p_figs = workdir / (docx.stem + '-figcheck.json')
    p_report = workdir / REPORT_NAME

    sig = docx_signature(docx, args.parts)
    old = None
    if p_report.exists():
        try:
            old = load_json(p_report)
        except Exception:
            old = None
    up_to_date = bool(old and old.get('signature') == sig and not args.force)
    print(f'文档: {docx}')
    print(f'工作目录: {workdir}')
    print(f'幂等状态: {"已是最新（可跳过已完成步骤）" if up_to_date else ("强制重跑" if args.force else "首次/文档已变更，全部重建")}')

    steps = []
    fail_step = None

    def record(name, status, detail):
        steps.append({'name': name, 'status': status, **detail})

    def hard_fail(name, msg):
        nonlocal fail_step
        fail_step = name
        print(f'\n✗ 步骤 [{name}] 失败：{msg}')

    # ---- 1. extract ----
    if up_to_date and p_paras.exists():
        record('extract', 'skip', {'out': str(p_paras)})
        print('[extract] 已存在，跳过')
    else:
        ok, d = run_step('extract', [TOOL_EXTRACT, docx, p_paras], cwd=workdir)
        record('extract', 'run' if ok else 'fail', {'out': str(p_paras), **d})
        if not ok or not p_paras.exists():
            hard_fail('extract', '提取正文段落失败（docx 损坏/加密/被占用？）')

    paras, lint_by_id = [], {}
    if not fail_step:
        try:
            paras = load_json(p_paras)
        except Exception as e:
            hard_fail('extract', f'paragraphs.json 读不出来: {e}')
    if not fail_step and not isinstance(paras, list):
        hard_fail('extract', 'paragraphs.json 应为数组')
    if not fail_step and not paras:
        hard_fail('extract', '未提取到任何正文段落（该文档可能全是标题/表格/图片）')
    if not fail_step and args.parts > len(paras):
        print(f'  ! --parts {args.parts} > 段落数 {len(paras)}，自动降为 {len(paras)} 片')

    # ---- 2. lint ----
    if not fail_step:
        if up_to_date and p_lint.exists():
            record('lint', 'skip', {'out': str(p_lint)})
            print('[lint] 已存在，跳过')
        else:
            ok, d = run_step('lint', [TOOL_LINT, p_paras, p_lint], cwd=workdir)
            record('lint', 'run' if ok else 'fail', {'out': str(p_lint), **d})
            if not ok:
                if args.lenient:
                    print('  ! --lenient：lint 失败按告警处理，继续')
                else:
                    hard_fail('lint', 'lint_paragraphs.py 失败')

    # ---- 3. check_refs ----
    if not fail_step:
        if args.skip_refs:
            record('refs', 'skip-option', {'out': ''})
            print('[refs] --skip-refs，跳过')
        elif up_to_date and p_refs.exists():
            record('refs', 'skip', {'out': str(p_refs)})
            print('[refs] 已存在，跳过')
        else:
            ok, d = run_step('refs', [TOOL_REFS, p_paras, '-o', p_refs], cwd=workdir)
            record('refs', 'run' if ok else 'fail', {'out': str(p_refs), **d})
            if not ok and not args.lenient:
                hard_fail('refs', 'check_refs.py 失败（对照表路径可用 --laws/--stds 手工重跑）')
            elif not ok:
                print('  ! --lenient：引用核查失败按告警处理，继续')

    # ---- 4. check_figures_tables ----
    if not fail_step:
        if args.skip_figures:
            record('figures', 'skip-option', {'out': ''})
            print('[figures] --skip-figures，跳过')
        elif up_to_date and p_figs.exists():
            record('figures', 'skip', {'out': str(p_figs)})
            print('[figures] 已存在，跳过')
        else:
            ok, d = run_step('figures', [TOOL_FIGS, docx, '-o', p_figs], cwd=workdir)
            record('figures', 'run' if ok else 'fail', {'out': str(p_figs), **d})
            if not ok and not args.lenient:
                hard_fail('figures', 'check_figures_tables.py 失败')
            elif not ok:
                print('  ! --lenient：图注核查失败按告警处理，继续')

    # ---- 5. 分片 ----
    part_stats = []
    if not fail_step:
        if p_lint.exists():
            try:
                lint_by_id = {x['id']: x for x in load_json(p_lint)}
            except Exception:
                lint_by_id = {}
        old_stats = (old or {}).get('part_stats') or []
        can_skip = (not args.force and old_stats
                    and len(old_stats) == args.parts
                    and sum(s['paras'] for s in old_stats) == len(paras)
                    and sum(s['chars'] for s in old_stats) == sum(len(p['text']) for p in paras)
                    and all((workdir / f'paragraphs_part{s["part"]}.json').exists()
                            for s in old_stats))
        if can_skip:
            record('split', 'skip', {'parts': args.parts})
            part_stats = old_stats
            print(f'[split] 已存在 {args.parts} 片且统计一致，跳过')
        else:
            try:
                t0 = time.time()
                stats = split_parts(paras, lint_by_id, args.parts)
                manifest = write_parts(workdir, stats)
                print(f'[split] {len(stats)} 片（{time.time() - t0:.1f}s）')
                for line in manifest.strip().splitlines()[1:]:
                    print('    ' + line)
                part_stats = [{k: v for k, v in s.items() if k not in ('items', 'lint')}
                              for s in stats]
                record('split', 'run', {'parts': len(stats), 'out': str(workdir / 'parts-manifest.txt')})
            except Exception as e:
                hard_fail('split', f'分片失败: {e}')

    # ---- 6. 汇总 ----
    lint_sum, refs_sum, figs_sum = {}, {}, {}
    if not fail_step:
        try:
            lint_sum = summarize_lint(p_lint) if p_lint.exists() else {
                'file': '', 'issues': 0, 'paragraphs_hit': 0, 'by_rule': {}}
            refs_sum = summarize_refs(p_refs) if p_refs.exists() else {
                'file': '', 'issues': 0, 'by_status': {}}
            figs_sum = summarize_figs(p_figs) if p_figs.exists() else {
                'file': '', 'issues': 0, 'by_category': {}}
        except Exception as e:
            hard_fail('summarize', f'汇总中间产物失败: {e}')

    total_chars = sum(len(p['text']) for p in paras) if paras else 0
    conc = suggest_concurrency(len(part_stats) or args.parts)
    head, todo = [], []
    if not fail_step:
        head, todo = build_todo(docx, workdir, len(paras), total_chars,
                                len(part_stats) or args.parts, lint_sum, refs_sum, figs_sum, conc)

    report = {
        'tool': 'prep.py',
        'version': 1,
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'ok': fail_step is None,
        'failed_step': fail_step,
        'docx': str(docx),
        'workdir': str(workdir),
        'signature': sig,
        'options': {'parts': args.parts, 'force': args.force, 'lenient': args.lenient,
                    'skip_refs': args.skip_refs, 'skip_figures': args.skip_figures},
        'steps': steps,
        'paragraphs': {'file': str(p_paras), 'count': len(paras), 'chars': total_chars},
        'parts': len(part_stats) or args.parts,
        'part_stats': part_stats,
        'lint': lint_sum,
        'refs': refs_sum,
        'figures': figs_sum,
        'recommended_concurrency': conc,
        'summary': head,
        'todo': todo,
    }
    save_json(p_report, report)

    print('\n' + '=' * 60)
    if fail_step:
        print(f'备料失败：步骤 [{fail_step}]（详见 {p_report}）')
        sys.exit(1)
    print('【备料完成】' + docx.name)
    for line in head:
        print('  ' + line)
    print('  分片明细：')
    for s in part_stats:
        print(f'    part{s["part"]}: {s["paras"]} 段 / {s["chars"]} 字 / '
              f'id {s["id_first"]}..{s["id_last"]} / lint {s.get("lint_hits", 0)} 条')
    if lint_sum.get('by_rule'):
        print('  lint 命中类型: ' + ', '.join(f'{k}×{v}' for k, v in lint_sum['by_rule'].items()))
    if refs_sum.get('by_status'):
        print('  引用问题类型: ' + ', '.join(f'{k}×{v}' for k, v in refs_sum['by_status'].items()))
    if figs_sum.get('by_category'):
        print('  图注问题类型: ' + ', '.join(f'{k}×{v}' for k, v in figs_sum['by_category'].items()))
    print('\n【下一步待办】')
    for i, line in enumerate(todo, 1):
        print(f'  {i}. {line}')
    print(f'\n报告 → {p_report}')
    return 0


if __name__ == '__main__':
    sys.exit(main() or 0)
