# -*- coding: utf-8 -*-
"""
assemble_report.py — 将各组 subagent 的结构化结论与 requirements.json 合成为校审报告。

默认输出 HTML；--format md 输出 Markdown；--format both 两者都出：
  <文档名>_校审报告.html    仪表板式自包含 HTML（顶部横幅、侧边目录、KPI卡、筛选、表格化结论）—— 默认
  <文档名>_校审报告.md      Markdown
两种格式均正确渲染上标/下标（如 m3/s→m³/s、10的8次方→10⁸、km²、H2O→H₂O）。

输入：
  <run_dir>/findings.json   —— [{"category","results":[{id,name,verdict,conclusion,evidence:[{loc,quote}]}]}]
  <run_dir>/meta.json
  $SKILL/scripts/requirements.json
用法：python assemble_report.py <run_dir> [--format md|html|both]（默认 html）
"""

from __future__ import annotations
import html
import json
import os
import re
import sys

# GBK 控制台自愈：打印 ✓/✗ 等会 UnicodeEncodeError 且产物不落盘（2026-09-30 实测）
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass
from collections import OrderedDict

VERDICTS = ["符合", "部分符合", "不符合", "不适用", "未检出"]
VCLASS = {"符合": "v-pass", "部分符合": "v-partial",
          "不符合": "v-fail", "不适用": "v-na", "未检出": "v-miss"}
# 行级配色（表格行）
VROW = {"符合": "row-passed", "部分符合": "row-partial",
        "不符合": "row-issue", "不适用": "row-na", "未检出": "row-issue"}

CAT_TITLES = {
    "强制性条文检查": "一、强制性条文检查",
    "常见设计问题检查": "二、常见设计问题检查",
    "一致性检查": "三、一致性检查",
    "语法表述检查": "四、语法表述检查",
    "文字规范性检查": "五、文字规范性检查",
    "设计标准检查": "六、设计标准检查",
    "法律法规检查": "七、法律法规检查",
    "表格逻辑关系专项检查": "八、表格逻辑关系专项检查",
}


# ============================================================
# 上标 / 下标 渲染（m3/s→m³/s、10的8次方→10⁸、km²、H2O 等）
# 思路：先把各种上下标记法统一成内部标记 ^{...} / _{...}，再按格式转成 <sup>/<sub>。
# ============================================================
_USUP = {"⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5",
         "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9", "⁺": "+", "⁻": "-", "ⁿ": "n"}
_USUB = {"₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4", "₅": "5",
         "₆": "6", "₇": "7", "₈": "8", "₉": "9", "₊": "+", "₋": "-",
         "ₙ": "n", "ₐ": "a", "ₘ": "m", "ᵢ": "i", "ₜ": "t"}
_CHEM = {"H2O": "H_{2}O", "CO2": "CO_{2}", "SO2": "SO_{2}", "SO4": "SO_{4}",
         "NO2": "NO_{2}", "NO3": "NO_{3}", "NH3": "NH_{3}", "NH4": "NH_{4}",
         "PO4": "PO_{4}", "HCO3": "HCO_{3}", "BOD5": "BOD_{5}"}


def supsub_markers(text):
    """把上/下标的各种写法统一为内部标记 ^{...} / _{...}（保留在纯文本中，安全）。"""
    if not text:
        return text
    t = text
    t = re.sub(r"(\d+)\s*的\s*(\d+)\s*次方",
               lambda m: f"{m.group(1)}^{{{m.group(2)}}}", t)
    t = re.sub(r"(\d)\^\(([^)]+)\)", lambda m: f"{m.group(1)}^{{{m.group(2)}}}", t)
    t = re.sub(r"(\d)\^(\d+)", lambda m: f"{m.group(1)}^{{{m.group(2)}}}", t)
    t = re.sub(r"[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻ⁿ]+",
               lambda m: "^{" + "".join(_USUP[c] for c in m.group(0)) + "}", t)
    t = re.sub(r"[₀₁₂₃₄₅₆₇₈₉₊₋ₙₐₘᵢₜ]+",
               lambda m: "_{" + "".join(_USUB[c] for c in m.group(0)) + "}", t)
    t = re.sub(r"(?<![A-Za-z])([kmc]?m|dm)(2|3)(?![0-9A-Za-z])",
               lambda m: f"{m.group(1)}^{{{m.group(2)}}}", t)
    for k, v in _CHEM.items():
        t = t.replace(k, v)
    return t


def markers_to_tags(text):
    """内部标记 ^{...} / _{...} → <sup>/<sub>。"""
    text = re.sub(r"\^\{([^}]*)\}", r"<sup>\1</sup>", text)
    text = re.sub(r"_\{([^}]*)\}", r"<sub>\1</sub>", text)
    return text


def fmt_html(raw):
    return markers_to_tags(html.escape(supsub_markers(raw)))


def fmt_md(raw):
    return markers_to_tags(supsub_markers(raw))


# ============================================================
# 定位转换：P0154→第16页、T039→表3.7.2.2（用 locators.json）
# ============================================================
_PCRE = re.compile(r"(?<![A-Za-z0-9])P(\d{4})(?!\d)")
_TCRE_A = re.compile(r"表T0?(\d{1,3})(?!\d)")
_TCRE_B = re.compile(r"(?<![A-Za-z0-9])T0?(\d{1,3})(?!\d)")


def make_xform(locators):
    """返回一个函数：把文本中的 P编号/T编号 转成用户可定位的'第N页'/'表X.X.X'。"""
    page = (locators or {}).get("page_by_para", {})
    heading = (locators or {}).get("heading_by_para", {})
    tables = (locators or {}).get("tables", {})

    def psub(m):
        n = int(m.group(1))
        pg = page.get(str(n))
        if pg:
            return f"第{pg}页"
        hd = heading.get(str(n))
        return f"§{hd}" if hd else m.group(0)

    def tsub(m):
        cap = tables.get(str(int(m.group(1))), {}).get("caption_num", "")
        return f"表{cap}" if cap else m.group(0)

    def xform(text):
        if not text:
            return text
        text = _TCRE_A.sub(tsub, text)   # 先处理 '表T039'
        text = _TCRE_B.sub(tsub, text)   # 再处理独立 'T039'
        text = _PCRE.sub(psub, text)     # 'P0154'→第N页
        return text

    return xform


def verdict_of(by_id, rid):
    res = by_id.get(rid)
    return (res.get("verdict", "未检出") if res else "未检出"), res


def _normalize_result_fields(r):
    """类型安全：evidence 必须为 list；issues 必须为 list[dict]。"""
    ev = r.get("evidence", [])
    if isinstance(ev, str):
        ev = []
    r["evidence"] = ev
    iss = r.get("issues", [])
    if not isinstance(iss, list):
        iss = []
    r["issues"] = [i for i in iss if isinstance(i, dict)]
    return r


def load(run_dir, here):
    reqs = json.load(open(os.path.join(here, "requirements.json"), encoding="utf-8"))
    meta = json.load(open(os.path.join(run_dir, "meta.json"), encoding="utf-8"))
    fp = os.path.join(run_dir, "findings.json")
    raw = json.load(open(fp, encoding="utf-8")) if os.path.exists(fp) else []
    by_id = {}
    for grp in raw:
        for r in grp.get("results", []):
            _normalize_result_fields(r)
            by_id[r["id"]] = r

    # 双保险：det_findings.json 的同 id 结果覆盖 findings.json（det 优先）。
    # 正常流程下 collect_findings 已合并，此处只在 collect 未跑或 findings 缺失时兜底。
    det_fp = os.path.join(run_dir, "det_findings.json")
    if os.path.exists(det_fp):
        det_raw = json.load(open(det_fp, encoding="utf-8"))
        n_det = 0
        for grp in det_raw:
            for r in grp.get("results", []):
                _normalize_result_fields(r)
                # 给 det 来源的条目在 conclusion 标注（机器判定），便于人工复核时识别
                if r.get("_source") == "deterministic" or r["id"] in by_id:
                    conc = r.get("conclusion", "")
                    if "（机器判定）" not in conc:
                        r["conclusion"] = conc + "　（机器判定）" if conc else "（机器判定）"
                by_id[r["id"]] = r   # det 覆盖
                n_det += 1
        if n_det:
            print(f"[assemble_report] det_findings 双保险并入 {n_det} 条（覆盖优先）")

    cats = OrderedDict()
    for r in reqs:
        cats.setdefault(r["category"], []).append(r)
    doc_name = os.path.splitext(meta.get("file", "报告"))[0]
    doc_dir = os.path.dirname(meta.get("docx_path", run_dir))
    lp = os.path.join(run_dir, "locators.json")
    locators = json.load(open(lp, encoding="utf-8")) if os.path.exists(lp) else {}
    return reqs, by_id, meta, cats, doc_name, doc_dir, locators


# ============================================================
# Markdown 渲染
# ============================================================
def render_markdown(reqs, by_id, meta, cats, doc_name, xform=None, max_issues=30):
    xf = xform or (lambda t: t)
    lines = []
    lines.append(f"# 《{doc_name}》智能校审报告\n")
    lines.append(f"- 文件：{meta.get('file','')}　段落：{meta.get('paragraphs','-')}　"
                 f"表格：{meta.get('tables','-')}　标题：{len(meta.get('headings',[]))}\n")
    lines.append("- 校审依据：检查要求.md（8 大类 42 条）；标准清单：水利标准.md；法规清单：法律法规.md\n")
    lines.append("- 生成方式：确定性脚本（抽取+机器证据）+ 并行语义检查（subagent）\n")
    lines.append("\n> 每条检查均附【要求原文】【结论】【依据/定位】；即便符合也给出主要分析结论。\n")
    for cat, items in cats.items():
        lines.append(f"\n## {CAT_TITLES.get(cat, cat)}\n")
        for req in items:
            verdict, res = verdict_of(by_id, req["id"])
            conclusion = (res.get("conclusion", "").strip() if res else "（未取得该条检查结论，建议复核。）")
            evidence = (res.get("evidence", []) if res else [])
            lines.append(f"### {req['id']} {req['name']}\n")
            lines.append(f"> **要求**：{fmt_md(req['content'])}\n")
            lines.append(f"**结论**：{verdict}\n")
            lines.append(f"**分析**：{fmt_md(xf(conclusion))}\n")
            issues = (res.get("issues") if res else None) or []
            if issues:
                lines.append("**具体问题与建议**：")
                for idx, it in enumerate(issues[:max_issues], 1):
                    loc = xf(it.get("loc", ""))
                    org = fmt_md(xf(it.get("original", "")))
                    typ = xf(it.get("type", ""))
                    sug = fmt_md(xf(it.get("suggestion", "")))
                    lines.append(f"{idx}. `{loc}` 原文「{org}」【{typ}】→ 建议：{sug}")
                if len(issues) > max_issues:
                    lines.append(f"……（共 {len(issues)} 条，其余 {len(issues) - max_issues} 条未显示；可用 --max-issues 调大）")
                lines.append("")
            if evidence:
                lines.append("**依据**：")
                for ev in evidence[:6]:
                    loc = xf(ev.get("loc", ""))
                    quote = (ev.get("quote", "") or "").strip().replace("\n", " ")
                    lines.append(f"- `{loc}`「{fmt_md(xf(quote[:160]))}」" if quote else f"- `{loc}`")
                lines.append("")
            elif not issues:
                lines.append("**依据**：（无）\n")
    lines.append("\n---\n## 校审结果汇总\n")
    lines.append("| 类别 | 条数 | " + " | ".join(VERDICTS) + " |")
    lines.append("|---|---|" + "|".join(["---"] * len(VERDICTS)) + "|")
    total = {k: 0 for k in VERDICTS}
    for cat, items in cats.items():
        row = {k: 0 for k in VERDICTS}
        for req in items:
            v, _ = verdict_of(by_id, req["id"])
            row[v] = row.get(v, 0) + 1
            total[v] += 1
        lines.append(f"| {cat} | {len(items)} | " + " | ".join(str(row.get(k, 0)) for k in VERDICTS) + " |")
    lines.append(f"| **合计** | **{len(reqs)}** | " + " | ".join(f"**{total.get(k,0)}**" for k in VERDICTS) + " |")
    lines.append("")
    return "\n".join(lines)


# ============================================================
# HTML 渲染（仪表板式：横幅 + 侧边目录 + KPI卡 + 筛选 + 表格化结论）
# ============================================================
HTML_STYLE = """
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:"Microsoft YaHei","PingFang SC","Helvetica Neue",Arial,sans-serif;
       line-height:1.6; color:#333; background:#f5f7fa; }
.header { background:linear-gradient(135deg,#6c5ce7,#a855f7,#7c3aed); color:#fff;
          padding:36px 60px; box-shadow:0 4px 20px rgba(108,92,231,.3); }
.header h1 { font-size:26px; font-weight:700; margin-bottom:6px; }
.header .meta { font-size:13.5px; opacity:.9; margin-top:10px; }
.header .meta span { margin-right:22px; display:inline-block; }
.container { max-width:none; margin:0; padding:24px 20px 60px 14px; display:flex; gap:20px; }
.sidebar { width:230px; flex-shrink:0; position:sticky; top:18px;
           max-height:calc(100vh - 36px); overflow-y:auto; background:#fff;
           border-radius:12px; padding:14px; box-shadow:0 2px 12px rgba(0,0,0,.06); }
.sidebar h3 { font-size:13px; color:#888; margin-bottom:10px; padding-bottom:8px;
              border-bottom:1px solid #eee; }
.sidebar ul { list-style:none; }
.sidebar ul li { margin-bottom:3px; }
.sidebar ul li a { display:block; padding:6px 10px; font-size:13px; color:#555;
                   text-decoration:none; border-radius:6px; white-space:nowrap;
                   overflow:hidden; text-overflow:ellipsis; transition:all .2s; }
.sidebar ul li a:hover { background:#f0eeff; color:#6c5ce7; }
.sidebar ul li a.active { background:#6c5ce7; color:#fff; }
.main-content { flex:1; min-width:0; max-width:1600px; margin:0 auto; }
.stats-cards { display:flex; gap:14px; margin-bottom:18px; flex-wrap:wrap; }
.stat-card { flex:1; min-width:120px; background:#fff; border-radius:12px; padding:18px;
             text-align:center; box-shadow:0 2px 12px rgba(0,0,0,.06); }
.stat-card .number { font-size:30px; font-weight:700; margin-bottom:2px; }
.stat-card .label { font-size:12.5px; color:#888; }
.stat-card.total .number { color:#6c5ce7; }
.stat-card.passed .number { color:#10b981; }
.stat-card.issues .number { color:#ef4444; }
.stat-card.rate .number { color:#3b82f6; }
.stat-card.na .number { color:#9ca3af; }
.filter-bar { background:#fff; border-radius:12px; padding:10px 18px; margin-bottom:18px;
              box-shadow:0 2px 12px rgba(0,0,0,.06); display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
.filter-bar .label { font-size:13px; color:#888; margin-right:6px; }
.filter-btn { padding:5px 15px; border:1px solid #ddd; border-radius:20px; background:#fff;
              color:#555; font-size:13px; cursor:pointer; transition:all .2s; }
.filter-btn:hover { border-color:#6c5ce7; color:#6c5ce7; }
.filter-btn.active { background:#6c5ce7; color:#fff; border-color:#6c5ce7; }
.report-body { background:#fff; border-radius:12px; padding:26px 34px;
               box-shadow:0 2px 12px rgba(0,0,0,.06); }
.report-body > p.intro { font-size:13.5px; color:#666; margin-bottom:8px; }
.report-body h2 { font-size:19px; color:#1a1a2e; margin:26px 0 12px; padding-bottom:6px;
                  border-bottom:2px solid #6c5ce7; scroll-margin-top:18px; }
.review-table { width:100%; border-collapse:collapse; table-layout:fixed; margin:0 0 8px; font-size:13.5px; }
.review-table th { background:#f8f7ff; color:#6c5ce7; padding:9px 11px; text-align:left;
                   border-bottom:2px solid #e5e0ff; font-weight:600; white-space:nowrap; }
.review-table td { padding:10px 11px; border-bottom:1px solid #f0f0f0; vertical-align:top; overflow-wrap:break-word; }
.review-table tr:hover td { background:#fafbff; }
.review-table td.req { color:#555; font-size:13px; }
.review-table td .rid { font-size:11px; color:#aaa; margin-top:2px; font-family:ui-monospace,Consolas,monospace; }
.review-table tr.row-passed td { color:#8a8a8a; }
.review-table tr.row-na td { color:#9aa0a6; background:#fafbfc; }
.review-table tr.row-partial td { background:#fffaf2; }
.review-table tr.row-partial:hover td { background:#fff3e0; }
.review-table tr.row-issue td { background:#fef5f5; }
.review-table tr.row-issue:hover td { background:#fde8e8; }
/* 有问题的内容（不符合/部分符合）：检查结果列加粗，结论(analysis)红色粗体 */
.review-table tr.row-issue td:nth-child(3),
.review-table tr.row-partial td:nth-child(3) { font-weight:700; }
.review-table tr.row-issue td:last-child .analysis,
.review-table tr.row-partial td:last-child .analysis { color:#b91c1c; font-weight:700; }
/* 具体问题清单（原文→建议） */
.issues { margin:7px 0 0; padding-left:22px; font-weight:400; color:#444; font-size:13px; }
.issues li { margin:5px 0; line-height:1.6; }
.issues li.imore { list-style:none; color:#9aa0a6; font-size:12px; }
.issues .iloc { font-family:ui-monospace,Consolas,monospace; background:#eef1f5; padding:1px 6px;
                border-radius:4px; font-size:11.5px; color:#666; }
.issues .iorg { color:#c62828; font-weight:700; }   /* 原文(问题) 红色粗体 */
.issues .itype { color:#9aa0a6; font-size:11.5px; margin:0 4px; }
.issues .isug { color:#1b8a3a; }                     /* 建议 绿色 */
.badge { font-size:12px; font-weight:600; padding:2px 10px; border-radius:999px; white-space:nowrap; }
.v-pass { background:#e8f5e9; color:#1b8a3a; border:1px solid #c8e6c9; }
.v-partial { background:#fff3e0; color:#e65100; border:1px solid #ffe0b2; }
.v-fail { background:#ffebee; color:#c62828; border:1px solid #ffcdd2; }
.v-na { background:#f0f0f0; color:#666; border:1px solid #e3e3e3; }
.v-miss { background:#f3e5f5; color:#6a1b9a; border:1px solid #e1bee7; }
.ev { margin-top:6px; padding-top:5px; border-top:1px dashed #eee; font-size:12.5px; color:#777; font-weight:400; }
.ev .evttl { color:#999; margin-bottom:2px; }
.ev div { margin:2px 0; line-height:1.55; }
.ev code, .rid { background:#eef1f5; padding:1px 6px; border-radius:4px; }
.footer { text-align:center; color:#9aa0a6; font-size:12px; margin:24px 0 4px; }
@media print {
  body { background:#fff; }
  .header { padding:18px 28px; box-shadow:none; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  .sidebar, .filter-bar { display:none; }
  .container { padding:0; }
  .stat-card { box-shadow:none; border:1px solid #ddd; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  .report-body { box-shadow:none; padding:10px 18px; }
  .review-table tr.row-issue td, .review-table tr.row-partial td,
  .review-table tr.row-na td { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
}
"""

HTML_SCRIPT = """
<script>
// 平滑滚动 + 目录高亮
document.querySelectorAll('#toc-list a').forEach(function(a){
  a.addEventListener('click', function(e){
    e.preventDefault();
    var tgt = document.querySelector(this.getAttribute('href'));
    if (tgt) tgt.scrollIntoView({behavior:'smooth', block:'start'});
  });
});
var headers = document.querySelectorAll('.report-body h2');
var links = document.querySelectorAll('#toc-list a');
window.addEventListener('scroll', function(){
  var cur = '';
  headers.forEach(function(h){ if (h.getBoundingClientRect().top <= 120) cur = h.id; });
  links.forEach(function(a){
    a.classList.toggle('active', a.getAttribute('href') === '#' + cur);
  });
});
// 筛选：all / issues(部分符合+不符合) / fail(不符合)
function filterResults(mode, btn){
  document.querySelectorAll('.filter-btn').forEach(function(b){ b.classList.remove('active'); });
  btn.classList.add('active');
  document.querySelectorAll('.review-table tbody tr').forEach(function(tr){
    var show = (mode === 'all')
      || (mode === 'issues' && (tr.classList.contains('row-partial') || tr.classList.contains('row-issue')))
      || (mode === 'fail' && tr.classList.contains('row-issue'));
    tr.style.display = show ? '' : 'none';
  });
  // 隐藏没有可见行的分类标题与其表格
  document.querySelectorAll('.report-body h2').forEach(function(h2){
    var tbl = h2.nextElementSibling;
    if (!tbl || tbl.tagName !== 'TABLE') return;
    var vis = tbl.querySelectorAll('tbody tr').length;
    var any = Array.from(tbl.querySelectorAll('tbody tr')).some(function(r){ return r.style.display !== 'none'; });
    h2.style.display = any ? '' : 'none';
    tbl.style.display = any ? '' : 'none';
  });
}
</script>
"""


def _stat_card(cls, number, label):
    return (f"<div class='stat-card {cls}'><div class='number'>{number}</div>"
            f"<div class='label'>{label}</div></div>")


def render_html(reqs, by_id, meta, cats, doc_name, xform=None, max_issues=30):
    esc = html.escape
    xf = xform or (lambda t: t)
    total = {k: 0 for k in VERDICTS}
    for req in reqs:
        v, _ = verdict_of(by_id, req["id"])
        total[v] += 1
    n_total = len(reqs)
    n_pass = total["符合"]
    n_issue = total["部分符合"] + total["不符合"]
    rate = f"{n_pass / n_total * 100:.1f}%" if n_total else "0%"
    n_na = total["不适用"]

    P = []
    P.append("<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>")
    P.append("<meta name='viewport' content='width=device-width, initial-scale=1.0'>")
    P.append(f"<title>校审报告 - {esc(doc_name)}</title>")
    P.append(f"<style>{HTML_STYLE}</style></head><body>")

    # 顶部横幅
    P.append("<div class='header'>")
    P.append(f"<h1>智能校审报告 — {esc(doc_name)}</h1>")
    P.append("<div class='meta'>")
    P.append(f"<span>文件：{esc(meta.get('file', ''))}</span>"
             f"<span>段落：{meta.get('paragraphs', '-')}</span>"
             f"<span>表格：{meta.get('tables', '-')}</span>"
             f"<span>检查项：{n_total}</span>")
    P.append("<span>校审依据：检查要求.md（8 大类 42 条）· 水利标准.md · 法律法规.md</span>")
    P.append("</div></div>")

    P.append("<div class='container'>")
    # 侧边目录
    P.append("<aside class='sidebar' id='toc'><h3>目录导航</h3><ul id='toc-list'>")
    for ci, cat in enumerate(cats.keys(), 1):
        P.append(f"<li><a href='#cat-{ci}'>{esc(cat)}</a></li>")
    P.append("</ul></aside>")

    # 主内容
    P.append("<div class='main-content'>")
    # KPI 卡
    P.append("<div class='stats-cards'>")
    P.append(_stat_card("total", n_total, "检查项总数"))
    P.append(_stat_card("passed", n_pass, "符合"))
    P.append(_stat_card("issues", n_issue, "存在问题"))
    P.append(_stat_card("rate", rate, "符合率"))
    if n_na:
        P.append(_stat_card("na", n_na, "不适用"))
    P.append("</div>")
    # 筛选条
    P.append("<div class='filter-bar'><span class='label'>筛选：</span>")
    P.append("<button class='filter-btn active' onclick=\"filterResults('all',this)\">全部</button>")
    P.append("<button class='filter-btn' onclick=\"filterResults('issues',this)\">仅存在问题（部分符合+不符合）</button>")
    P.append("<button class='filter-btn' onclick=\"filterResults('fail',this)\">仅不符合</button>")
    P.append("</div>")

    # 报告正文（表格化）
    P.append("<div class='report-body'>")
    P.append("<p class='intro'>每条检查均附【要求】【检查结果】【说明与建议】；点击右上筛选可只看问题项；"
             "左侧目录可快速跳转。结论色标："
             "<span class='badge v-pass'>符合</span> "
             "<span class='badge v-partial'>部分符合</span> "
             "<span class='badge v-fail'>不符合</span> "
             "<span class='badge v-na'>不适用</span></p>")

    for ci, (cat, items) in enumerate(cats.items(), 1):
        title = CAT_TITLES.get(cat, cat)
        P.append(f"<h2 id='cat-{ci}'>{esc(title)}</h2>")
        P.append("<table class='review-table'><thead><tr>"
                 "<th style='width:12%'>检查项</th>"
                 "<th style='width:20%'>要求</th>"
                 "<th style='width:10%'>检查结果</th>"
                 "<th style='width:58%'>说明 / 建议</th></tr></thead><tbody>")
        for req in items:
            verdict, res = verdict_of(by_id, req["id"])
            rowclass = VROW.get(verdict, "row-na")
            vclass = VCLASS.get(verdict, "v-na")
            conclusion = (res.get("conclusion", "") if res
                          else "（未取得该条检查结论，建议复核。）")
            issues = (res.get("issues") if res else None) or []
            # 说明/建议：结论 + 具体问题清单(原文→建议)
            cell = [f"<div class='analysis'>{fmt_html(xf(conclusion))}</div>"]
            if issues:
                cell.append("<ol class='issues'>")
                for it in issues[:max_issues]:
                    loc = esc(xf(it.get("loc", "")))
                    org = fmt_html(xf(it.get("original", "")))
                    typ = esc(xf(it.get("type", "")))
                    sug = fmt_html(xf(it.get("suggestion", "")))
                    item = (f"<li><span class='iloc'>{loc}</span> "
                            f"<span class='iorg'>原文「{org}」</span>")
                    if typ:
                        item += f"<span class='itype'>[{typ}]</span>"
                    item += f"<br><span class='isug'>建议：{sug}</span></li>"
                    cell.append(item)
                if len(issues) > max_issues:
                    cell.append(f"<li class='imore'>……共 {len(issues)} 条，其余 {len(issues) - max_issues} 条未显示（可用 --max-issues 调大）</li>")
                cell.append("</ol>")
            desc = "".join(cell)
            P.append(
                f"<tr class='{rowclass}'>"
                f"<td><strong>{esc(req['name'])}</strong><div class='rid'>{esc(req['id'])}</div></td>"
                f"<td class='req'>{fmt_html(req['content'])}</td>"
                f"<td><span class='badge {vclass}'>{esc(verdict)}</span></td>"
                f"<td>{desc}</td></tr>")
        P.append("</tbody></table>")

    P.append(f"<div class='footer'>由「报告智能校审」技能生成 · 共 {n_total} 条检查"
             f"（符合 {n_pass} · 部分符合 {total['部分符合']} · 不符合 {total['不符合']} · 不适用 {n_na}）</div>")
    P.append("</div>")  # report-body
    P.append("</div>")  # main-content
    P.append("</div>")  # container
    P.append(HTML_SCRIPT)
    P.append("</body></html>")
    return "\n".join(P)


def main():
    # 兼容 "--format md" 与 "--format=md" 两种写法（历史上只认等号，空格写法静默回落默认 html）
    argv = sys.argv[1:]
    normalized = []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ("--format", "--max-issues") and i + 1 < len(argv) and not argv[i + 1].startswith("--"):
            normalized.append(a + "=" + argv[i + 1])
            i += 2
        else:
            normalized.append(a)
            i += 1
    args = [a for a in normalized if not a.startswith("--")]
    # 默认输出 HTML；用户指定 --format md 时输出 Markdown；--format both 两者都出
    fmt = "html"
    max_issues = 30
    for a in normalized:
        if a.startswith("--format="):
            fmt = a.split("=", 1)[1]
        elif a.startswith("--max-issues="):
            try:
                max_issues = int(a.split("=", 1)[1])
            except ValueError:
                pass
    run_dir = args[0] if args else None
    if not run_dir:
        print("用法: python assemble_report.py <run_dir> [--format md|html|both]（默认 html）")
        sys.exit(1)
    here = os.path.dirname(os.path.abspath(__file__))
    reqs, by_id, meta, cats, doc_name, doc_dir, locators = load(run_dir, here)
    xform = make_xform(locators)

    written = []
    if fmt in ("md", "both"):
        md = render_markdown(reqs, by_id, meta, cats, doc_name, xform, max_issues)
        p = os.path.join(doc_dir, f"{doc_name}_校审报告.md")
        open(p, "w", encoding="utf-8").write(md)
        written.append(p)
    if fmt in ("html", "both"):
        h = render_html(reqs, by_id, meta, cats, doc_name, xform, max_issues)
        p = os.path.join(doc_dir, f"{doc_name}_校审报告.html")
        open(p, "w", encoding="utf-8").write(h)
        written.append(p)
    dist = {k: 0 for k in VERDICTS}
    for req in reqs:
        v, _ = verdict_of(by_id, req["id"])
        dist[v] += 1
    print(f"[assemble_report] 覆盖 {len(reqs)} 条；分布 {dist}")
    for p in written:
        print(f"[assemble_report] 已生成 {p}")


if __name__ == "__main__":
    main()
