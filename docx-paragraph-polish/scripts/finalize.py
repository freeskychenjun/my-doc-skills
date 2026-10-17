# -*- coding: utf-8 -*-
"""finalize.py —— 父代理收口一条龙（固化版 finalize3.py）。

每轮都要现写的"提交脚本"（实测泸水第三轮被写了 3 次全文 + 15 次编辑，占父代理
输出 token 的 23.9%）。**通用收口逻辑全部固化在本脚本**，项目特有的内容只走配置：

  通用（本脚本，不改动）：
    · 标点/体例统一（全角括号、～、引号、文号空格、错别字映射表、用词表）
    · 保护清单（OMML/上下标格式/编制依据清单/公式行/技术附录路径）
    · 双重上标防护（单位上标由 run 级格式实现时还原成普通数字）
    · 数字 guard（两道：改写后 & 体例统一后；多重集口径 ²/³ 先归一）
    · 相似度 guard（改写幅度 < 0.5 回退）
    · 批注清理与定稿（CLEAR_FLAGS / FLAG_FINAL 探针）
  项目特有（配置文件 finalize_config.json）：
    · typo / term            本项目错别字与用词收口映射
    · reffix / reffix_note   查证遍定向修正（id → [[旧串, 新串], ...]）+ 定稿批注文案
    · num_exempt             允许改变数字多重集的段（如"删去重复年份引用"）
    · flag_final             定稿批注：id → {"probes": [特征串], "text": "批注"}
    · clear_flags            经查证属误报、清空批注的段 id

用法:
  python finalize.py <paragraphs.json> [--merged merged.json] [-o final.json] \
      --src <docx> [--fmt-scan format_scan.json] [--config finalize_config.json] \
      [--protect-out protected_ids.json] [--dry]

  --src 与 --fmt-scan 二选一：给了 --fmt-scan（scan_protect.py 产物）就不必再读 docx；
  都没有时按"无格式防护"降级并在报告里告警。
  退出码: 0 正常（含 --dry）；1 参数/读文件错误。

设计约束（与 P6 的 A4 一致）：本脚本**不含任何某一轮的项目内容**，换项目只换配置；
首轮若出现全新的 guard 逻辑，应作为本脚本的新选项/新配置项落地，而不是再抄一份脚本。
"""
import argparse
import json
import re
import sys
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from lint_paragraphs import auto_fix          # noqa: E402
from scan_protect import build_protected, scan_formatted  # noqa: E402

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

NUM = re.compile(r'\d+(?:\.\d+)?')
WD = re.compile(r'〔(\d{4})〕\s*(\d+)\s+号')
SUPDIGITS = '⁰¹²³⁴⁵⁶⁷⁸⁹'
SUPMAP = str.maketrans(SUPDIGITS, '0123456789')
CJK = re.compile(r'[\u4e00-\u9fff]')


def is_cjk(ch):
    return bool(ch) and bool(CJK.match(ch))


def norm_quotes(text):
    out = list(text)
    for k, ch in enumerate(out):
        if ch != '"':
            continue
        before = out[k - 1] if k else ''
        after = out[k + 1] if k + 1 < len(out) else ''
        opened = 0
        for m in out[:k]:
            if m == '\u201c':
                opened += 1
            elif m == '\u201d':
                opened -= 1
        if opened > 0:
            out[k] = '\u201d'
        elif is_cjk(after):
            out[k] = '\u201c'
        elif is_cjk(before):
            out[k] = '\u201d'
    return ''.join(out)


def norm_parens(text):
    stack, pairs = [], []
    for k, ch in enumerate(text):
        if ch == '(':
            stack.append(k)
        elif ch == ')' and stack:
            pairs.append((stack.pop(), k))
    if '(' not in text and ')' not in text:
        return text
    out = list(text)
    for i, j in pairs:
        content = text[i + 1:j]
        before = text[i - 1] if i else ''
        after = text[j + 1] if j + 1 < len(text) else ''
        if CJK.search(content) or is_cjk(before) or is_cjk(after):
            out[i] = '\uff08'
            out[j] = '\uff09'
    return ''.join(out)


def norm_tilde(text):
    return re.sub(r'(?<=[\u4e00-\u9fff0-9)\uff09])~(?=[\u4e00-\u9fff0-9(\uff08])',
                  '\uff5e', text)


def numbag(text):
    """数字 token 多重集（²/³ 先归一；单位指数写法 km2→km 视为同一写法）。"""
    t = text.translate(SUPMAP)
    t = re.sub(r'(km|hm|cm|mm|m)\s*([23])(?=[^0-9]|$)', r'\1', t)
    return Counter(NUM.findall(t))


def load_config(path):
    cfg = {'typo': {}, 'term': {}, 'reffix': {}, 'reffix_note': {},
           'num_exempt': [], 'flag_final': {}, 'clear_flags': [], 'name': ''}
    if not path:
        return cfg
    raw = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    cfg.update({k: raw.get(k, cfg[k]) for k in cfg})
    cfg['reffix'] = {int(k): [tuple(x) for x in v] for k, v in cfg['reffix'].items()}
    cfg['reffix_note'] = {int(k): v for k, v in cfg['reffix_note'].items()}
    cfg['num_exempt'] = set(int(x) for x in cfg['num_exempt'])
    cfg['clear_flags'] = set(int(x) for x in cfg['clear_flags'])
    ff = {}
    for k, v in cfg['flag_final'].items():
        probes, text = v['probes'], v['text']
        ff[int(k)] = (list(probes), text)
    cfg['flag_final'] = ff
    return cfg


def make_normalizer(cfg):
    typo, term = cfg['typo'], cfg['term']

    def normalize(text):
        t = auto_fix(text)[0]
        for wrong, right in typo.items():
            t = t.replace(wrong, right)
        for wrong, right in term.items():
            t = t.replace(wrong, right)
        t = WD.sub(r'〔\1〕\2号', t)
        t = norm_tilde(t)
        t = norm_quotes(t)
        t = norm_parens(t)
        return t
    return normalize


def diagnose(paras, merged, fmt, protected, cfg):
    """只读诊断：逐段给出"会被哪道 guard 回退"的明细（供 check_guard.py / 收尾报告）。

    返回 {'rows': [{id, reason, detail}], 'counts': {reason: n}}。
    判据与 main() 里的 guard 链**同源**（同一 numbag / 同一 protect / 同一 REFFIX 豁免）。
    """
    normalize = make_normalizer(cfg)
    REFFIX, NUM_EXEMPT = cfg['reffix'], cfg['num_exempt']
    rows = []
    counts = Counter()
    for pid, p in paras.items():
        orig = p['text']
        m = merged.get(pid)
        base_text = m['newText'] if m else orig
        if pid in protected:
            if normalize(orig) != orig:
                rows.append({'id': pid, 'reason': '保护段-仅体例改动',
                             'detail': protected[pid]})
                counts['保护段-仅体例改动'] += 1
            continue
        if base_text == orig:
            continue
        info = fmt.get(p['paraIndex']) or {}
        if numbag(base_text) != numbag(orig) and pid not in NUM_EXEMPT and pid not in REFFIX:
            rows.append({'id': pid, 'reason': '数字guard回退',
                         'detail': 'orig=%s new=%s' % (
                             sorted(numbag(orig).elements())[:12],
                             sorted(numbag(base_text).elements())[:12])})
            counts['数字guard回退'] += 1
            continue
        if SequenceMatcher(None, orig, base_text).ratio() < 0.5:
            rows.append({'id': pid, 'reason': '相似度guard回退',
                         'detail': 'sim=%.3f' % SequenceMatcher(None, orig, base_text).ratio()})
            counts['相似度guard回退'] += 1
            continue
        bt = base_text
        for new, old in info.get('reverts', []):
            if new in bt and old not in bt:
                bt = bt.replace(new, old)
        bt = normalize(bt)
        for old_s, new_s in REFFIX.get(pid, []):
            if old_s in bt:
                bt = bt.replace(old_s, new_s)
        if pid not in NUM_EXEMPT and pid not in REFFIX and numbag(bt) != numbag(orig):
            rows.append({'id': pid, 'reason': '修正后数字异常回退',
                         'detail': '体例/查证修正后数字多重集不一致'})
            counts['修正后数字异常回退'] += 1
    return {'rows': rows, 'counts': dict(counts)}


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('paragraphs')
    ap.add_argument('--merged', default='')
    ap.add_argument('-o', '--out', default='')
    ap.add_argument('--src', default='')
    ap.add_argument('--fmt-scan', default='')
    ap.add_argument('--config', default='')
    ap.add_argument('--protect-out', default='')
    ap.add_argument('--extra-pattern', default='')
    ap.add_argument('--dry', action='store_true')
    args = ap.parse_args()

    pj = Path(args.paragraphs).resolve()
    if not pj.exists():
        print(f'paragraphs 不存在: {pj}')
        sys.exit(1)
    base = pj.parent
    paras = {p['id']: p for p in json.loads(pj.read_text(encoding='utf-8-sig'))}
    cfg = load_config(args.config)
    normalize = make_normalizer(cfg)
    REFFIX, REF_NOTE = cfg['reffix'], cfg['reffix_note']
    NUM_EXEMPT, FLAG_FINAL, CLEAR_FLAGS = cfg['num_exempt'], cfg['flag_final'], cfg['clear_flags']

    # ---- 格式扫描（OMML / 上下标格式 / 单位上标还原规则）----
    if args.fmt_scan:
        raw = json.loads(Path(args.fmt_scan).read_text(encoding='utf-8-sig'))
        fmt = {int(k): v for k, v in raw.items()}
        fmt_src = args.fmt_scan
    elif args.src:
        fmt = scan_formatted(Path(args.src).resolve())
        fmt_src = args.src
    else:
        fmt = {}
        fmt_src = '(未提供，无格式防护)'
        print('⚠ 未给 --src / --fmt-scan：跳过 OMML/上下标格式防护（建议先跑 scan_protect.py）')
    protected = build_protected(list(paras.values()), fmt, args.extra_pattern or None)

    merged = {}
    if args.merged:
        mp = Path(args.merged).resolve()
        if mp.exists():
            merged = {m['id']: m for m in json.loads(mp.read_text(encoding='utf-8-sig'))}
        else:
            print(f'⚠ merged 不存在（{mp}），按"无改写"处理')
    else:
        for cand in ('merged.json', 'final.json'):
            if (base / cand).exists():
                merged = {m['id']: m for m in
                          json.loads((base / cand).read_text(encoding='utf-8-sig'))}
                print(f'· 未指定 --merged，自动使用 {cand}')
                break

    # 定稿批注：配置没写文案、但 reffix 有 note 的，自动按"特征串 → 文案"补齐
    for pid, note in REF_NOTE.items():
        if pid in REFFIX and pid not in FLAG_FINAL:
            FLAG_FINAL[pid] = ([REFFIX[pid][0][0]], note)

    stats = Counter()
    out = []
    for pid, p in paras.items():
        orig = p['text']
        m = merged.get(pid)
        base_text = m['newText'] if m else orig
        flag = (m or {}).get('flag', '')
        info = fmt.get(p['paraIndex']) or {}

        if pid in protected:
            base_text = normalize(orig)
            if base_text != orig:
                stats['保护段-仅体例改动'] += 1
        else:
            if base_text != orig:
                if numbag(base_text) != numbag(orig) \
                        and pid not in NUM_EXEMPT and pid not in REFFIX:
                    base_text = normalize(orig)
                    flag = (flag + '；' if flag else '') + \
                        '父代理收口：该段改写导致数字与原文不一致，已回退原文。'
                    stats['数字guard回退'] += 1
                elif SequenceMatcher(None, orig, base_text).ratio() < 0.5:
                    base_text = normalize(orig)
                    flag = (flag + '；' if flag else '') + \
                        '父代理收口：该段改写幅度过大，已回退原文。'
                    stats['相似度guard回退'] += 1
                else:
                    stats['模型改写保留'] += 1
            for new, old in info.get('reverts', []):
                if new in base_text and old not in base_text:
                    base_text = base_text.replace(new, old)
                    stats['双重上标防护'] += 1
            base_text = normalize(base_text)
            if numbag(base_text) != numbag(orig) and pid not in NUM_EXEMPT:
                base_text = normalize(orig)
                flag = (flag + '；' if flag else '') + \
                    '父代理收口：体例统一后数字校验异常，已回退原文。'
                stats['体例后数字异常回退'] += 1
            for old_s, new_s in REFFIX.get(pid, []):
                if old_s in base_text:
                    base_text = base_text.replace(old_s, new_s)
                    stats['父代理查证修正'] += 1
            if pid in REFFIX:
                stats['查证修正（豁免数字guard）'] += 1
                if pid in REF_NOTE and REF_NOTE[pid] not in flag:
                    flag = (flag + '；' if flag else '') + REF_NOTE[pid]
                    stats['查证批注'] += 1
            elif pid in NUM_EXEMPT:
                if numbag(base_text) != numbag(orig):
                    stats['有意删除重复年份-豁免数字guard'] += 1
            elif numbag(base_text) != numbag(orig):
                base_text = normalize(orig)
                flag = (flag + '；' if flag else '') + \
                    '父代理收口：修正后数字校验异常，已回退原文。'
                stats['修正后数字异常回退'] += 1

        if base_text != orig:
            stats['最终有改动'] += 1
            if normalize(orig) == base_text and normalize(orig) != orig:
                stats['其中仅体例统一'] += 1
        if pid in CLEAR_FLAGS:
            flag = ''
            stats['已核实误报-清除批注'] += 1
        fin = FLAG_FINAL.get(pid)
        if fin:
            probes, text = fin
            if any(s in orig for s in probes):
                flag = text
                stats['定稿批注'] += 1
        entry = {'id': pid, 'paraIndex': p['paraIndex'], 'original': orig, 'newText': base_text}
        if flag:
            entry['flag'] = flag
        out.append(entry)

    out.sort(key=lambda x: x['paraIndex'])
    print('格式扫描来源: %s' % fmt_src)
    print('保护段 %d：' % len(protected))
    for k, v in Counter(protected.values()).most_common():
        print('   %s: %d' % (k, v))
    print('格式上标防护段 %d（reverts 就位）' %
          sum(1 for v in fmt.values() if v.get('reverts')))
    for k in ('模型改写保留', '保护段-仅体例改动', '双重上标防护', '数字guard回退',
              '相似度guard回退', '体例后数字异常回退', '父代理查证修正',
              '修正后数字异常回退', '最终有改动', '其中仅体例统一',
              '已核实误报-清除批注', '定稿批注'):
        print('  %s: %d' % (k, stats[k]))
    nflag = sum(1 for e in out if e.get('flag'))
    print('  批注段数: %d' % nflag)
    if args.dry:
        print('(dry run，未写盘)')
        return 0
    out_path = Path(args.out) if args.out else base / 'final.json'
    out_path.write_text(json.dumps(out, ensure_ascii=False), encoding='utf-8')
    prot_path = Path(args.protect_out) if args.protect_out else base / 'protected_ids.json'
    prot_path.write_text(json.dumps(protected, ensure_ascii=False, indent=1), encoding='utf-8')
    print('→ %s' % out_path)
    print('→ %s' % prot_path)
    return 0


if __name__ == '__main__':
    sys.exit(main() or 0)
