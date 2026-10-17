# -*- coding: utf-8 -*-
"""check_guard.py —— 收口 guard 的**只读诊断**：哪些段会被回退、为什么。

以前每轮现写（`check_guard.py` 在运行目录里被反复手抄），判据还容易与 finalize 跑偏。
本脚本直接复用 finalize.py 的 guard 链（同一 numbag / 同一保护清单 / 同一查证豁免），
所以它报告的"会被回退"与 finalize --dry 的统计**同源**。

典型用途：
  1. 收口前先跑一遍，确认"回退计数 == 0"，避免最后一步才发现数字 guard 静默回退；
  2. 调试"为什么某段没改成"——打印每段的数字多重集差异 / 相似度 / 保护原因；
  3. 写进 postapply_check.py，作为确定性终检的一环。

用法:
  python check_guard.py <paragraphs.json> [--merged merged.json] [--src docx] \
      [--fmt-scan format_scan.json] [--config finalize_config.json] \
      [--protect protected_ids.json] [--json out.json] [--fail-on-revert] [-v]

退出码: 0 = 无回退；3 = 有回退（告警，默认不阻塞）；加 --fail-on-revert 时 1。
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import finalize as FZ  # noqa: E402

sys.stdout.reconfigure(encoding='utf-8', errors='replace')


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('paragraphs')
    ap.add_argument('--merged', default='')
    ap.add_argument('--src', default='')
    ap.add_argument('--fmt-scan', default='')
    ap.add_argument('--config', default='')
    ap.add_argument('--protect', default='')
    ap.add_argument('--json', default='')
    ap.add_argument('--fail-on-revert', action='store_true')
    ap.add_argument('-v', '--verbose', action='store_true')
    args = ap.parse_args()

    pj = Path(args.paragraphs).resolve()
    if not pj.exists():
        print(f'paragraphs 不存在: {pj}')
        sys.exit(1)
    base = pj.parent
    paras = {p['id']: p for p in json.loads(pj.read_text(encoding='utf-8-sig'))}
    cfg = FZ.load_config(args.config)

    merged = {}
    mp = Path(args.merged) if args.merged else base / 'merged.json'
    if mp.exists():
        merged = {m['id']: m for m in json.loads(mp.read_text(encoding='utf-8-sig'))}
    else:
        print(f'⚠ merged 不存在: {mp}（将按"无改写"诊断，回退必然为 0）')

    if args.fmt_scan:
        fmt = {int(k): v for k, v in
               json.loads(Path(args.fmt_scan).read_text(encoding='utf-8-sig')).items()}
    elif args.src:
        from scan_protect import scan_formatted
        fmt = scan_formatted(Path(args.src).resolve())
    else:
        fmt = {}

    if args.protect:
        prot = {int(k): str(v) for k, v in
                json.loads(Path(args.protect).read_text(encoding='utf-8-sig')).items()}
    else:
        from scan_protect import build_protected
        prot = build_protected(list(paras.values()), fmt)

    rep = FZ.diagnose(paras, merged, fmt, prot, cfg)
    n_revert = sum(v for k, v in rep['counts'].items() if k != '保护段-仅体例改动')
    print('段落 %d；merged 有改写 %d 段；保护清单 %d 段'
          % (len(paras), sum(1 for k, m in merged.items()
                             if m.get('newText') != paras[k]['text']) if merged else 0,
             len(prot)))
    for k, v in sorted(rep['counts'].items(), key=lambda x: -x[1]):
        print('  %-16s %d' % (k + ':', v))
    if rep['rows']:
        print('\n明细（最多 40 条）：')
        for r in rep['rows'][:40]:
            print('  id=%-5s %-14s %s' % (r['id'], r['reason'], r['detail'][:150]))
        if len(rep['rows']) > 40:
            print('  …共 %d 条（--json 取全量）' % len(rep['rows']))
    if args.json:
        Path(args.json).write_text(json.dumps(rep, ensure_ascii=False, indent=1),
                                   encoding='utf-8')
        print('→ %s' % args.json)
    if n_revert:
        print('\n结论: 有 %d 段会被 guard 回退（dry-run 回退计数 ≠ 0，先查原因再出件）'
              % n_revert)
        sys.exit(1 if args.fail_on_revert else 3)
    print('\n结论: 无 guard 回退（dry-run 回退计数 = 0）')
    return 0


if __name__ == '__main__':
    sys.exit(main() or 0)
