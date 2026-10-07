# -*- coding: utf-8 -*-
"""提取 docx 中适合"校正"的正文段落（只读，不依赖 Word/WPS 安装）。

规则移植自 WordTools/SmartWriting/BodyParagraphFilter.cs：
排除 标题（大纲级别/标题样式/编号/章节模式）、表格内段落、图名/表名、
含图片/形状/域代码的段落、空段落。

用法:
  python extract_body_paragraphs.py <docx路径> [输出json路径]
默认输出 <docx同目录>/<文件名>-paragraphs.json
"""
import json
import locale
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

CAPTION_RE = re.compile(r'^[图表]\s*[\d一二三四五六七八九十]')
NUMBERED_HEADING_RE = re.compile(r'^(\d+(?:\.\d+)*)\s')
CHAPTER_RE = re.compile(r'^((第[一二三四五六七八九十百千\d]+)[章节篇部])')
HEADING_STYLE_RE = re.compile(r'Heading\s*\d', re.IGNORECASE)
MAX_HEADING_LENGTH = 60
MAX_CAPTION_LENGTH = 40

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def paragraph_text(p):
    texts = [t.text or '' for t in p.findall(f'.//{W}t')]
    return ''.join(texts)


def contains_special(p):
    """含图片/形状/域代码（修订写入会破坏或无意义）。"""
    for tag in ('drawing', 'pict', 'fldChar', 'instrText', 'object'):
        if p.findall(f'.//{W}{tag}'):
            return True
    return False


def is_heading(p, text):
    # 大纲级别 1-9（w:outlineLvl val=0..8；9=正文）
    ol = p.find(f'{W}pPr/{W}outlineLvl')
    if ol is not None:
        try:
            if 0 <= int(ol.get(f'{W}val')) <= 8:
                return True
        except (TypeError, ValueError):
            pass
    # 标题样式
    ps = p.find(f'{W}pPr/{W}pStyle')
    if ps is not None:
        name = ps.get(f'{W}val') or ''
        if '标题' in name or HEADING_STYLE_RE.search(name) or name.lower().startswith('heading'):
            return True
    if len(text) > MAX_HEADING_LENGTH:
        return False
    if NUMBERED_HEADING_RE.match(text):
        return True
    if CHAPTER_RE.match(text):
        return True
    return False


def is_body_paragraph(p):
    text = paragraph_text(p).strip()
    if not text:
        return False
    if contains_special(p):
        return False
    if len(text) <= MAX_CAPTION_LENGTH and CAPTION_RE.match(text):
        return False
    if is_heading(p, text):
        return False
    return True


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    doc_path = Path(sys.argv[1]).resolve()
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else doc_path.with_name(
        doc_path.stem + '-paragraphs.json')

    doc = Document(str(doc_path))
    body = doc.element.body

    items = []
    para_index = -1
    for child in body:
        if child.tag != f'{W}p':
            continue
        para_index += 1
        if not is_body_paragraph(child):
            continue
        text = paragraph_text(child).strip()
        items.append({
            'id': len(items) + 1,
            'paraIndex': para_index,
            'text': text,
        })

    out_path.write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding='utf-8')
    print(f'OK paragraphs={len(items)} out={out_path}')
    if not items:
        print('未找到可处理的正文段落（标题、表格内、图名表名、图片、域等均已跳过）。')


if __name__ == '__main__':
    main()
