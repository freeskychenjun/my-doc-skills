# -*- coding: utf-8 -*-
"""make_prompts.py —— 一次生成全部子代理任务书（pass1 / pass2 / 仲裁 / 查证遍）。

背景：以前每轮要临时写脚本把模板占位符替换掉（实测四个环节各手写一次，约 4 分钟编排开销，
且极易漏填/填错——`{{PART}}` 是 `{{PARTS}}` 的前缀，用字符串 replace 会互相污染）。
本脚本把这些固化成一条命令，并做「零残留占位符」校验。

用法:
  python make_prompts.py all    --workdir <运行目录> --report "<报告名>" [--round 第一轮]
  python make_prompts.py pass   --workdir <运行目录> --report "<报告名>" [--parts 6]
  python make_prompts.py arb    --workdir <运行目录> --report "<报告名>" [--parts 3]
  python make_prompts.py verify --workdir <运行目录> --report "<报告名>" [--batch-size 10]

常用可选项:
  --outdir DIR        任务书输出目录（默认 <workdir>/prompts）
  --background-file F 本轮背景（markdown 文件，原样注入 {{ROUND_BACKGROUND}}）
  --watch-file F      本报告专项留意（markdown 文件，原样注入 {{PROJECT_WATCH}}）
  --background TEXT / --watch TEXT   同上，直接给文本
  --dry               只校验输入是否齐备与占位符是否可闭合，不写文件

各子命令依赖的前置产物（由 prep.py / build_arb.py / collect_flags.py 产生）:
  pass   : paragraphs_part{1..N}.json + lint_part{1..N}.json
  arb    : arb_input_part{1..N}.json
  verify : flags.json（collect_flags.py 输出）+ paragraphs.json

输出:
  <outdir>/PROMPT_pass1_partN.md / PROMPT_pass2_partN.md / PROMPT_arb_partN.md / PROMPT_verify_batchN.md
  <outdir>/prompts-manifest.json   机器可读清单（每个任务书：类型/批次/条目数/绝对路径）

退出码: 0 全部生成并通过校验；1 缺前置产物 / 占位符未闭合 / 模板缺失。
"""
import argparse
import json
import re
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent
TPL_DIR = SKILL_DIR / 'templates'

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

PH = re.compile(r'\{\{([A-Z_]+)\}\}')
DRY = False          # --dry：只校验，不写盘（write_prompt 里判断）

# 查证遍按「证据源」路由分批：同一条证据链路的 flag 放一批，子代理不用来回换查证手段
VERIFY_ROUTES = [
    (('法规', '标准'), '法规·标准引用'),
    (('数字',), '文内一致性·数字'),
    (('地名', '其他'), '地名·术语·其他事实'),
]

DEFAULT_BG_FIRST = ('本报告为**首次校正（{r}）**：原稿未经本流程处理，没有上一轮基线，'
                    '不要假设某些段落已被润色过；按当前原文实际情况逐段判断。')
DEFAULT_BG_RETRY = ('本报告为**{r}复核**：上一轮的修订已被人工接受/拒绝，本轮只针对仍存在问题的段落；'
                    '已经规范的段落不要为改而改。')
DEFAULT_WATCH = ('## 本报告专项留意（父代理预置）\n\n'
                 '本轮无预置专项留意项。若你在段内发现跨段/全篇性问题（术语不统一、编号体系矛盾等），'
                 '**不要自行裁决**，写进 `flag` 交给第 6 步查证遍。')
ARB_DEFAULT_BG = ('**背景**：本报告已完成两遍独立改写。仲裁对象只有两类——'
                  '「两遍都改但文本不同」的冲突段，以及「只有一遍改过」的低置信段；'
                  '其余段落已由投票决定，不在你的范围内。')


def die(msg):
    print(f'✗ {msg}', file=sys.stderr)
    sys.exit(1)


def load_template(name: str) -> str:
    p = TPL_DIR / name
    if not p.exists():
        die(f'模板缺失: {p}')
    return p.read_text(encoding='utf-8')


def fill(tpl: str, mapping: dict, where: str) -> str:
    """用正则整体替换，避免 {{PART}} 吃掉 {{PARTS}} 的前缀（历史事故）。"""
    def repl(m):
        key = m.group(1)
        if key not in mapping:
            die(f'{where}: 模板占位符 {{{{{key}}}}} 没有对应取值（内部错误）')
        return str(mapping[key])
    out = PH.sub(repl, tpl)
    left = PH.findall(out)
    if left:
        die(f'{where}: 替换后仍有残留占位符 {sorted(set(left))}')
    return out


def write_prompt(outdir: Path, name: str, text: str, manifest: list, **meta):
    outdir.mkdir(parents=True, exist_ok=True)
    p = outdir / name
    if not DRY:
        p.write_text(text, encoding='utf-8')
    manifest.append({'file': str(p), 'chars': len(text), **meta})
    print(f'  · {name}  {len(text)} 字符' + (f'  items={meta["items"]}' if meta.get('items') else ''))
    return p


def round_background(args) -> str:
    if args.background_file:
        return Path(args.background_file).read_text(encoding='utf-8').strip()
    if args.background:
        return args.background
    r = args.round.replace('第', '').replace('轮', '').strip() or '第一轮'
    first = r in ('1', '一', '第一', '首次', '第一轮')
    tpl = DEFAULT_BG_FIRST if first else DEFAULT_BG_RETRY
    return tpl.format(r='第一轮' if first else args.round)


def project_watch(args) -> str:
    if args.watch_file:
        return Path(args.watch_file).read_text(encoding='utf-8').strip()
    if args.watch:
        return args.watch
    return DEFAULT_WATCH


def need(path: Path, hint: str):
    if not path.exists():
        die(f'缺前置产物 {path}\n  → {hint}')


# ---------------------------------------------------------------- pass
def cmd_pass(args, manifest):
    wd = Path(args.workdir)
    print(f'[pass] 生成 pass1 + pass2 任务书（{args.parts} 片）')
    for i in range(1, args.parts + 1):
        need(wd / f'paragraphs_part{i}.json', '先跑 scripts/prep.py')
        need(wd / f'lint_part{i}.json', '先跑 scripts/prep.py')
    bg, watch = round_background(args), project_watch(args)
    tpl1 = load_template('PROMPT_pass1.md')
    tpl2 = load_template('PROMPT_pass2.md')
    for pass_no, tpl in ((1, tpl1), (2, tpl2)):
        for i in range(1, args.parts + 1):
            m = {
                'REPORT_NAME': args.report, 'ROUND': args.round,
                'PARTS': args.parts, 'PART': i,
                'PART_FILE': str(wd / f'paragraphs_part{i}.json'),
                'LINT_FILE': str(wd / f'lint_part{i}.json'),
                'OUT_FILE': str(wd / f'pass{pass_no}_part{i}.json'),
                'ROUND_BACKGROUND': bg, 'PROJECT_WATCH': watch,
            }
            text = fill(tpl, m, f'PROMPT_pass{pass_no}_part{i}')
            write_prompt(args.outdir, f'PROMPT_pass{pass_no}_part{i}.md', text, manifest,
                         kind=f'pass{pass_no}', part=i)


# ---------------------------------------------------------------- arb
def cmd_arb(args, manifest):
    wd = Path(args.workdir)
    print(f'[arb] 生成仲裁任务书（{args.arb_parts} 包）')
    tpl = load_template('PROMPT_arb.md')
    bg = ARB_DEFAULT_BG
    if args.background_file:
        bg = Path(args.background_file).read_text(encoding='utf-8').strip()
    elif args.background:
        bg = args.background
    watch = project_watch(args)
    for i in range(1, args.arb_parts + 1):
        f = wd / f'arb_input_part{i}.json'
        need(f, '先跑 scripts/build_arb.py')
        data = json.loads(f.read_text(encoding='utf-8'))
        items = data if isinstance(data, list) else data.get('items', [])
        if not items:
            print(f'  ! arb_input_part{i}.json 为空，跳过')
            continue
        m = {
            'REPORT_NAME': args.report, 'ROUND': args.round,
            'PARTS': args.arb_parts, 'PART': i,
            'ARB_IN_FILE': str(f), 'ARB_OUT_FILE': str(wd / f'arb_part{i}.json'),
            'ARB_COUNT': len(items),
            'ROUND_BACKGROUND': bg, 'PROJECT_WATCH': watch,
        }
        text = fill(tpl, m, f'PROMPT_arb_part{i}')
        write_prompt(args.outdir, f'PROMPT_arb_part{i}.md', text, manifest,
                     kind='arb', part=i, items=len(items))


# ---------------------------------------------------------------- verify
def cmd_verify(args, manifest):
    wd = Path(args.workdir)
    flags_path = wd / 'flags.json'
    paras_path = wd / 'paragraphs.json'
    need(flags_path, '先跑 scripts/collect_flags.py')
    need(paras_path, '先跑 scripts/prep.py')
    data = json.loads(flags_path.read_text(encoding='utf-8'))
    items = data.get('items', []) if isinstance(data, dict) else data
    if not items:
        print('[verify] flags.json 里没有 flag，无需查证遍')
        return

    # 按证据源分组；同一组超过 batch_size 就再切分
    by_type = {}
    for it in items:
        by_type.setdefault(it.get('type') or '其他', []).append(it)
    groups, seen = [], set()
    for types, title in VERIFY_ROUTES:
        bucket = [x for t in types for x in by_type.get(t, [])]
        if not bucket:
            continue
        for k in range(0, len(bucket), args.batch_size):
            chunk = bucket[k:k + args.batch_size]
            suffix = '' if len(bucket) <= args.batch_size else f'（第 {k // args.batch_size + 1} 组）'
            groups.append((title + suffix, chunk))
            seen |= {x['id'] for x in chunk}
    rest = [x for x in items if x['id'] not in seen]      # 兜底：未归类的类型
    if rest:
        for k in range(0, len(rest), args.batch_size):
            groups.append(('未归类', rest[k:k + args.batch_size]))

    print(f'[verify] 生成查证任务书：{len(items)} 条 flag → {len(groups)} 批')
    tpl = load_template('PROMPT_verify_batch.md')
    for bid, (title, chunk) in enumerate(groups, 1):
        payload = [{
            'id': it['id'],
            'paraIndex': it.get('paraIndex'),
            'type': it.get('type'),
            '原文': it.get('original'),
            '当前候选文本': it.get('newText'),
            '已有flag': it.get('flag'),
        } for it in chunk]
        m = {
            'REPORT_NAME': args.report,
            'BATCH_ID': bid, 'BATCH_TITLE': title, 'N': len(chunk),
            'FLAGS_JSON': json.dumps(payload, ensure_ascii=False, indent=1),
            'PARAS': str(paras_path),
            'OUT_FILE': str(wd / f'verify{bid}.json'),
        }
        text = fill(tpl, m, f'PROMPT_verify_batch{bid}')
        write_prompt(args.outdir, f'PROMPT_verify_batch{bid}.md', text, manifest,
                     kind='verify', batch=bid, title=title, items=len(chunk))
        print(f'      ids={[x["id"] for x in chunk]}')
    covered = {x['id'] for _, c in groups for x in c}
    missing = {x['id'] for x in items} - covered
    if missing:
        die(f'有 {len(missing)} 条 flag 未被任何批次覆盖: {sorted(missing)[:10]}')


def main():
    ap = argparse.ArgumentParser(
        description='一次生成 pass1/pass2/仲裁/查证遍全部任务书（模板占位符填充 + 零残留校验）')
    ap.add_argument('cmd', choices=['all', 'pass', 'arb', 'verify'])
    ap.add_argument('--workdir', required=True, help='运行目录（prep.py 的 --workdir）')
    ap.add_argument('--report', required=True, help='报告名（写进任务书）')
    ap.add_argument('--round', default='第一轮')
    ap.add_argument('--parts', type=int, default=6, help='pass=分片数')
    ap.add_argument('--arb-parts', type=int, default=3, help='仲裁包数（独立于分片数）')
    ap.add_argument('--batch-size', type=int, default=10, help='verify=每批条数上限')
    ap.add_argument('--outdir', default='', help='任务书输出目录，默认 <workdir>/prompts')
    ap.add_argument('--background', default='', help='本轮背景文本')
    ap.add_argument('--background-file', default='', help='本轮背景 markdown 文件')
    ap.add_argument('--watch', default='', help='本报告专项留意文本')
    ap.add_argument('--watch-file', default='', help='本报告专项留意 markdown 文件')
    ap.add_argument('--dry', action='store_true', help='只校验，不写文件')
    args = ap.parse_args()

    wd = Path(args.workdir).resolve()
    if not wd.is_dir():
        die(f'--workdir 不存在: {wd}')
    args.outdir = Path(args.outdir).resolve() if args.outdir else wd / 'prompts'

    manifest = []
    globals()['DRY'] = bool(args.dry)

    if args.cmd in ('all', 'pass'):
        cmd_pass(args, manifest)
    if args.cmd in ('all', 'arb'):
        try:
            cmd_arb(args, manifest)
        except SystemExit:
            if args.cmd != 'all':
                raise
            print('  （仲裁阶段缺 arb_input_part*.json，先跑 build_arb.py 再单独执行 make_prompts.py arb）')
    if args.cmd in ('all', 'verify'):
        try:
            cmd_verify(args, manifest)
        except SystemExit:
            if args.cmd != 'all':
                raise
            print('  （查证阶段缺 flags.json，先跑 collect_flags.py 再单独执行 make_prompts.py verify）')

    if not args.dry:
        mp = args.outdir / 'prompts-manifest.json'
        mp.write_text(json.dumps({'workdir': str(wd), 'report': args.report,
                                  'round': args.round, 'prompts': manifest},
                                 ensure_ascii=False, indent=1), encoding='utf-8')
        print(f'\n✓ 共 {len(manifest)} 份任务书 → {args.outdir}')
        print(f'  清单: {mp}')
        kinds = {}
        for m in manifest:
            kinds[m['kind']] = kinds.get(m['kind'], 0) + 1
        print('  分派: ' + '，'.join(f'{k}×{v}' for k, v in sorted(kinds.items()))
              + '（并发按限流表取值：编码套餐型 3–4，通用 API 型 6–8）')
    else:
        print(f'\n(dry) 可生成 {len(manifest)} 份任务书，未写盘')


if __name__ == '__main__':
    main()
