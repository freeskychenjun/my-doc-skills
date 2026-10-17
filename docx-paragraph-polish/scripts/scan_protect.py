# -*- coding: utf-8 -*-
"""OMML / 上下标格式 / 保护清单扫描（一条命令）。

每轮都需要、以前每轮现写（omml_scan.py + omml_scan2.py + 保护清单目测），本脚本固化：

1. 扫出台面 docx 的**正文顶层段落**里哪些段落在"重建段落"时有风险：
   - 含 OMML 公式对象（m:oMath / m:oMathPara）→ 重建会丢公式，必须保护；
   - 含**段内上下标格式**（run 级 vertAlign=superscript/subscript）→ 重建可能错位，必须保护；
   - 含单位上标由格式实现（如 `100km` + 上标 `2`，文本层看起来是 km100² 类）→
     记录"格式上标 → 普通数字"的 revert 规则，供 finalize 做双重上标防护。
2. 按固定口径生成**保护清单**（protected_ids.json，{id: 原因}）：
   - 公式段 / 含上下标格式段（上面两类）；
   - 编制依据/引用清单行（以 [n] 《…》 开头）；
   - 公式行（带编号 (1-1) 之类）；
   - 技术附录里含文件名/路径/配置键的行（PROTECT_PAT）；
   - 可加 `--extra-pattern` 补充项目特有的保护正则。
3. 输出 `format_scan.json`（逐 paraIndex 的 protect/reverts 明细）供 finalize.py --fmt-scan 复用，
   避免 finalize 再算一遍。

用法:
  python scan_protect.py <docx> <paragraphs.json> \
      [-o protected_ids.json] [--fmt-out format_scan.json] [--report scan_protect.txt] \
      [--extra-pattern "正则"] [--json]

退出码: 0 正常；1 参数/读文件错误。
"""
import argparse
import json
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
SUPDIGITS = '⁰¹²³⁴⁵⁶⁷⁸⁹'

# 技术附录/配置说明里不该改的行（文件名、盘符路径、配置键）
DEFAULT_PROTECT_PAT = (
    r'(\.exe|\.bat|\.CFG|\.INF|\.RAI|\.LEV|\.IC1D|\.IC2D|\.RTC|\.txt|\.csv'
    r'|Miniutes|run_GD|run_SD|River_dream|input_data|run_config|[A-Za-z]:\\'
    r'|后缀为|run_title|run_type|start_time|time_step|initial_1d|initial_2d'
    r'|文件夹)')
REF_LIST_PAT = re.compile(r'^\s*\[?\d{1,3}\]?\s*[\u300a]')
FORMULA_LINE_PAT = re.compile(r'^\s*[^\u4e00-\u9fff]{0,40}\(\d+[-–]\d+\)\s*$')
CJK = re.compile(r'[\u4e00-\u9fff]')


def scan_formatted(docx_path):
    """逐正文顶层段落扫描，返回 {paraIndex: {protect, reverts, why}}。"""
    from docx import Document
    doc = Document(str(docx_path))
    out = {}
    for i, p in enumerate([c for c in doc.element.body if c.tag == W + 'p']):
        if p.findall('.//' + M + 'oMath') or p.findall('.//' + M + 'oMathPara'):
            out[i] = {'protect': True, 'reverts': [], 'why': '含 OMML 公式对象'}
            continue
        acc = ''
        subs = False
        reverts = []
        formula_sup = False
        for r in p.iter(W + 'r'):
            txt = ''.join(t.text or '' for t in r.findall(W + 't'))
            rpr = r.find(W + 'rPr')
            va = rpr.find(W + 'vertAlign') if rpr is not None else None
            val = va.get(W + 'val') if va is not None else None
            if val == 'subscript':
                subs = True
            elif val == 'superscript' and txt:
                if all(c in '0123456789' for c in txt):
                    m = re.search(r'(\d[\d.,]*)(km|hm|m)$', acc)
                    m2 = re.search(r'([\d.,]+)×10$', acc)
                    if m:
                        for d in txt:
                            reverts.append(('%s%s%s' % (m.group(1), m.group(2), SUPDIGITS[int(d)]),
                                            '%s%s%s' % (m.group(1), m.group(2), d)))
                    elif m2:
                        for d in txt:
                            reverts.append(('%s×10%s' % (m2.group(1), SUPDIGITS[int(d)]),
                                            '%s×10%s' % (m2.group(1), d)))
                    else:
                        formula_sup = True
                else:
                    formula_sup = True
            acc += txt
        if subs or formula_sup:
            out[i] = {'protect': True, 'reverts': [], 'why': '含上下标格式的公式段'}
        elif reverts:
            out[i] = {'protect': False, 'reverts': reverts, 'why': '单位上标为格式实现'}
    return out


def build_protected(paras, fmt, extra_pat=None):
    prot = {}
    pat = re.compile(DEFAULT_PROTECT_PAT + (('|' + extra_pat) if extra_pat else ''), re.I)
    for p in paras:
        t, pid = p['text'], p['id']
        if pid in prot:
            continue
        info = fmt.get(p['paraIndex'])
        if info and info.get('protect'):
            prot[pid] = info['why']
        elif REF_LIST_PAT.match(t.strip()):
            prot[pid] = '编制依据/引用清单'
        elif FORMULA_LINE_PAT.match(t.strip()):
            prot[pid] = '公式行(带编号)'
        elif pat.search(t):
            prot[pid] = '技术附录含文件名/路径'
    return prot


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('docx')
    ap.add_argument('paragraphs')
    ap.add_argument('-o', '--out', default='')
    ap.add_argument('--fmt-out', default='')
    ap.add_argument('--report', default='')
    ap.add_argument('--extra-pattern', default='')
    ap.add_argument('--json', action='store_true', help='只打印机器可读摘要')
    args = ap.parse_args()

    docx = Path(args.docx).resolve()
    pj = Path(args.paragraphs).resolve()
    if not docx.exists() or not pj.exists():
        print(f'输入不存在: {docx} / {pj}')
        sys.exit(1)
    paras = json.loads(pj.read_text(encoding='utf-8-sig'))
    fmt = scan_formatted(docx)
    prot = build_protected(paras, fmt, args.extra_pattern or None)

    # 统计
    n_omml = sum(1 for v in fmt.values() if v.get('why') == '含 OMML 公式对象')
    n_sup = sum(1 for v in fmt.values() if v.get('why') == '含上下标格式的公式段')
    n_rev = sum(1 for v in fmt.values() if v.get('reverts'))
    by_reason = {}
    for r in prot.values():
        by_reason[r] = by_reason.get(r, 0) + 1
    risky = [p for p in paras if p['paraIndex'] in fmt]
    summary = {
        'docx': str(docx), 'paragraphs': len(paras),
        'omml_paras': n_omml, 'sup_format_paras': n_sup,
        'unit_superscript_reverts': n_rev,
        'protected': len(prot), 'by_reason': by_reason,
        'protected_in_rewrite_scope': len([p for p in risky if p['id'] in prot]),
    }

    out = Path(args.out) if args.out else pj.parent / 'protected_ids.json'
    out.write_text(json.dumps(prot, ensure_ascii=False, indent=1), encoding='utf-8')
    summary['protected_ids'] = str(out)
    if args.fmt_out:
        fo = Path(args.fmt_out)
        fo.write_text(json.dumps({str(k): v for k, v in fmt.items()},
                                 ensure_ascii=False, indent=1), encoding='utf-8')
        summary['format_scan'] = str(fo)

    lines = ['保护清单扫描：%s' % docx.name,
             '正文段落(提取) %d  → 保护 %d 段' % (len(paras), len(prot)),
             '  含 OMML 公式对象的段落 %d' % n_omml,
             '  含段内上下标格式的段落 %d' % n_sup,
             '  需做"格式上标→数字"还原的段落 %d' % n_rev]
    for k, v in sorted(by_reason.items(), key=lambda x: -x[1]):
        lines.append('  保护原因 %-22s %d' % (k + ':', v))
    lines.append('')
    lines.append('保护段 id 清单: %s' % sorted(prot)[:80] + (' …' if len(prot) > 80 else ''))
    if n_rev:
        lines.append('')
        lines.append('单位上标（格式实现）还原规则（前 20 条）：')
        shown = 0
        for idx, v in sorted(fmt.items()):
            for a, b in v.get('reverts', []):
                lines.append('  paraIndex=%s  %s → %s' % (idx, a, b))
                shown += 1
                if shown >= 20:
                    break
            if shown >= 20:
                break
    text = '\n'.join(lines)
    print(text)
    if args.report:
        Path(args.report).write_text(text, encoding='utf-8')
    if args.json:
        print('JSON ' + json.dumps(summary, ensure_ascii=False))
    print('→ %s' % out)


if __name__ == '__main__':
    main()
