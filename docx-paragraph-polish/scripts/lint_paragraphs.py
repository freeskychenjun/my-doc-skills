# -*- coding: utf-8 -*-
"""确定性预检：用错别字映射表 + 正则规则扫描正文段落，产出"确定性问题清单"。

这些机械问题（错别字/标点/单位写法/多余空格）不依赖模型判断，脚本直接扫出；
模型改写时拿着本清单只需"确认执行"，把注意力留给表达层问题。

用法:
  python lint_paragraphs.py <paragraphs.json> [输出lint.json]

输出: [{"id":1,"issues":[{"rule":"进人→进入","type":"错别字","pos":123,"snippet":"…进人平原…","fix":"进入"}]},…]
      + 控制台统计（按规则分组计数）
"""
import json
import re
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ---------- 规则表 ----------
# 1) 错别字/OCR 映射（直接替换级，确定无疑的）
TYPO_MAP = {
    '进人': '进入', '注人': '注入', '汇人': '汇入', '流人': '流入', '排人': '排入',
    '机率': '几率', '淘空': '掏空', '座落在': '坐落在', '汉道': '汊道', '湖汉': '湖汊',
    '最高记录': '最高纪录', '历史记录': '历史纪录', '实测记录': '实测纪录',
}

# 2) 正则规则：(规则名, 正则, 修复模板说明, 类型)
REGEX_RULES = [
    # 单位大小写（SI 应小写）
    ('Km→km', re.compile(r'(\d)\s*Km\b'), '小写 km', '单位写法'),
    ('KM→km', re.compile(r'(\d)\s*KM\b'), '小写 km', '单位写法'),
    # 数字与单位间多余空格
    ('数字单位间空格', re.compile(r'(\d)\s+(km|mm|cm|m3|m²|km2|km³|hm2|℃|‰|%|亩)(?![\dA-Za-z])'), '去空格', '空格'),
    # 文号/编号内空格（如 "40 号"）
    ('文号内空格', re.compile(r'〔\d{4}〕\s*\d+\s+号'), '去空格', '空格'),
    # 中文句中半角句号/逗号（句末英文句点）
    ('半角句点', re.compile(r'[\u4e00-\u9fff]\.(?=\s|$|[\u4e00-\u9fff])'), '全角。', '标点'),
    # 中文旁的半角括号
    ('半角括号', re.compile(r'[\u4e00-\u9fff]\(|\)[\u4e00-\u9fff]'), '全角（）', '标点'),
    # 中文旁的半角直引号
    ('半角引号', re.compile(r'[\u4e00-\u9fff]"|"[\u4e00-\u9fff]'), '全角“”', '标点'),
    # 中文旁的半角波浪号
    ('半角波浪号', re.compile(r'\d~\d'), '全角～', '标点'),
    # "1、2月份" 这类月份列举（顿号应为"1月、2月"）
    ('月份列举顿号', re.compile(r'([1-9]|1[0-2])、([1-9]|1[0-2])月份'), 'N月、N月', '标点'),
    # "X月X日"缺"日"（自6月8开始 → 6月8日）
    ('缺日字', re.compile(r'([1-9]|1[0-2])月([1-9]|[12]\d|3[01])(?![日号\s])((?:开始|结束|起|后|前|以来|之后|之前))'), '补"日"', '漏字'),
]

# 中文句中的半角逗号容易被误报（数字千分位等），单独一条谨慎规则：
# 仅当半角逗号两侧均为汉字时报
REGEX_RULES.append(('汉字间半角逗号', re.compile(r'[\u4e00-\u9fff],[\u4e00-\u9fff]'), '全角，', '标点'))


# ---------- 自动修复（零风险规则，可直接替换，不经模型） ----------
AUTO_REGEX = [
    ('Km→km', re.compile(r'(\d)\s*Km\b'), r'\1km'),
    ('KM→km', re.compile(r'(\d)\s*KM\b'), r'\1km'),
    ('数字单位间空格', re.compile(r'(\d)\s+(km2|km²|km3|km|mm|cm|m3|m²|hm2|℃|‰|%|亩)(?![\dA-Za-z])'), r'\1\2'),
    ('文号内空格', re.compile(r'〔(\d{4})〕\s*(\d+)\s+号'), r'〔\1〕\2号'),
]


def auto_fix(text):
    """应用零风险确定性修复（错别字映射 + 单位大小写 + 多余空格）。
    返回 (新文本, 应用的规则列表)。"""
    applied = []
    for wrong, right in TYPO_MAP.items():
        if wrong in text:
            text = text.replace(wrong, right)
            applied.append(f'{wrong}→{right}')
    for name, rx, repl in AUTO_REGEX:
        new = rx.sub(repl, text)
        if new != text:
            applied.append(name)
            text = new
    return text, applied


def lint(text):
    issues = []
    # 映射表
    for wrong, right in TYPO_MAP.items():
        start = 0
        while True:
            pos = text.find(wrong, start)
            if pos < 0:
                break
            issues.append({'rule': f'{wrong}→{right}', 'type': '错别字', 'pos': pos,
                           'snippet': text[max(0, pos - 8):pos + len(wrong) + 8], 'fix': right})
            start = pos + len(wrong)
    # 正则
    for item in REGEX_RULES:
        name, rx = item[0], item[1]
        fix = item[2] if len(item) > 2 else ''
        typ = item[3] if len(item) > 3 else '其他'
        for m in rx.finditer(text):
            issues.append({'rule': name, 'type': typ, 'pos': m.start(),
                           'snippet': text[max(0, m.start() - 8):m.end() + 8], 'fix': fix})
    issues.sort(key=lambda x: x['pos'])
    return issues


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else src.rsplit('.', 1)[0] + '-lint.json'
    paras = json.load(open(src, encoding='utf-8'))
    results = []
    n_issues = 0
    from collections import Counter
    by_rule = Counter()
    for p in paras:
        issues = lint(p['text'])
        if issues:
            results.append({'id': p['id'], 'paraIndex': p['paraIndex'], 'issues': issues})
            n_issues += len(issues)
            for i in issues:
                by_rule[i['rule']] += 1
    json.dump(results, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'OK 段落 {len(paras)}，命中段落 {len(results)}，问题 {n_issues} 条 → {out}')
    for rule, c in by_rule.most_common():
        print(f'  {rule}: {c}')


if __name__ == '__main__':
    main()
