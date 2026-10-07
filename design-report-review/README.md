# 报告智能校审 技能（skill 标识：`design-report-review`）

对水利工程设计报告（Word `.docx`/`.doc`）进行**逐条、可追溯**的智能校审，覆盖 8 大类 42 条检查要求，输出与要求一一对应、引用完整要求原文的校审报告。**默认输出仪表板式 HTML**（顶部横幅 + 左侧目录 + KPI 统计卡 + 问题筛选 + 按类别表格化结论、行级配色），用户指定时输出 Markdown；两种格式均正确渲染上标/下标（如 m³/s、10⁸、km²、H₂O）。

## 目录结构

```
design-report-review/
├── SKILL.md                      技能定义与编排流程（核心）
├── README.md                     本文件
├── 使用说明.html                 用户使用说明（HTML，浏览器打开）
├── references/                   参考/规则（可更新）
│   ├── 检查要求.md               8 大类 42 条检查项与输出要求（权威）
│   ├── 水利标准.md               水利标准现行有效清单
│   └── 法律法规.md               法律法规现行有效清单
├── scripts/
│   ├── lib_docx.py               docx 解析基础库（顺序遍历/合并单元格重建/正则/有效数字）
│   ├── build_reference_index.py  解析两份清单 → references_index.json
│   ├── extract_docx.py           docx → 结构化 artifacts（主抽取器）
│   ├── com_docx_extract.py       Word COM 内存直读器（.doc 受 DGS 拦截时用，绕过写盘）
│   ├── deterministic_checks.py   机器可验证证据（表格核算/引用对照/有效数字/编号）
│   ├── slice_artifacts.py        按分组裁剪产物 → _slices/groupA..D.json（提速核心）
│   ├── det_findings.py           确定性结论直出 std_001/003/004、gram_007（提质提速核心）
│   ├── collect_findings.py       收集 subagent 产物 + det_findings，合并/校验/全覆盖
│   ├── polish_findings.py        结论润色（去除机器字段名黑话，翻译为自然语言）
│   ├── assemble_report.py        合成 HTML / Markdown 报告（det 双保险）
│   ├── parse_requirements.py     解析检查要求.md → requirements.json
│   ├── page_map.py               用 Word COM 计算每段 Word 页码 → locators.json（可选）
│   ├── prepare.py                一键"抽取+确定性检查"（旧版分步，新流程见 SKILL.md）
│   ├── maintain.py               清单维护：一键重建索引/要求（智能/强制/检查/备份）
│   ├── references_index.json     (生成) 标准/法规检索索引
│   └── requirements.json         (生成) 42 条结构化要求
└── _run/<文档名>/                 (生成) 每次运行的 artifacts 与证据
    ├── fulltext.md tables.json captions.json citations.json numbers.json meta.json locators.json
    ├── det_table.json det_citations.json det_numbers.json det_numbering.json
    ├── det_findings.json         (生成) 4 条确定性直出结论
    ├── _slices/groupA..D.json    (生成) 按分组裁剪的切片（subagent 唯一输入）
    ├── groupA..D.json            (生成) 4 组 subagent 的结构化结论
    └── findings.json             (生成) 合并后的全部 42 条结论（合成报告来源）
```

## 8 大类 42 条检查项

| 大类 | 编号 | 条数 | 主要数据来源 |
|------|------|------|--------------|
| 强制性条文检查 | mp_001–011 | 11 | 全文(语义) |
| 常见设计问题检查 | ci_* | 9 | 全文 + 排涝模数数值 |
| 一致性检查 | cons_001–003 | 3 | 全文 + 表格 + 数值 |
| 语法表述检查 | gram_001–008 | 8 | 全文 + 题录 |
| 文字规范性检查 | std_001–004 | 4 | 机器证据 det_numbers |
| 设计标准检查 | ds_001–002 | 2 | 机器证据 det_citations + 索引 |
| 法律法规检查 | lr_001–002 | 2 | 机器证据 det_citations + 索引 |
| 表格逻辑关系专项检查 | tbl_001–003 | 3 | 表格网格 + det_table |

## 工作原理（确定性 + 语义 的混合，6 步流程）

1. **抽取**（`extract_docx.py` / `com_docx_extract.py`，确定性）：保留文档顺序，重建合并单元格，产出全文/表格/题录/引用/数值五类结构化 artifacts。`.doc` 自动转换；若环境有数据安全软件（如 Agile DGS）拦截 Word 写盘，改用 `com_docx_extract.py` 内存直读绕过。
2. **确定性证据**（`deterministic_checks.py`）：高精度产出机器可验证结论——表格统计值核算、标准/法规编号有效性对照清单、流量有效数字（按 SL/T 247 规范，整数末尾零不算有效数字）、工程等别/级别罗马-阿拉伯、图表编号顺序与格式。**宁可漏报(留给语义)也不误报**。
3. **切片 + 直出**（`slice_artifacts.py` + `det_findings.py`，提速提质核心）：按 4 组裁剪产物为小切片（每个 <2 万 token）；其中 4 条纯规则条目（std_001/003/004、gram_007）由 `det_findings.py` 直接生成结论，**不派 subagent**——这些是数数/查表/加法，脚本比 LLM 更准。
4. **并行语义检查**（Claude 按 SKILL.md fan-out 4 个 subagent）：A=强条+常见、B=一致性+表格、C=语法+文字、D=标准+法规。每个 subagent 只读对应切片，互不阻塞；半直出条目（tbl_001/ds/lr）直接引用切片内的 det 证据。
5. **收集 + 润色**（`collect_findings.py` + `polish_findings.py`）：合并 4 组 subagent 产物与 det 直出结论（det 优先），校验字段、确保 42 条全覆盖；再润色结论，把 subagent 偶尔照抄的机器字段名翻译成自然语言。
6. **合成**（`assemble_report.py`，确定性）：以 `requirements.json` 为序遍历全部 42 条，确保一一对应、引用完整要求原文；det 来源条目标注"（机器判定）"。

## 使用方式

### 作为 Claude Code 技能（推荐）
将本目录置于 `.claude/skills/design-report-review/`，随后在对话中：
> 用报告智能校审技能检查 `安徽两湖涝区项建第2、3、5章.docx`

Claude 会读取 `SKILL.md` 并自动执行上述流程，**默认**在同目录生成 `<文档名>_校审报告.html`；若用户要求 Markdown，则生成 `<文档名>_校审报告.md`。

完整的用户使用说明（检查范围 / 输出报告样例 / 清单动态维护）见 [`使用说明.html`](使用说明.html)，浏览器直接打开。

### 手动分步（便于调试/复现，新 6 步流程）
```bash
cd .claude/skills/design-report-review
# 0) 建索引与要求（仅首次或清单更新后）
python scripts/maintain.py --force   # 一键重建 references_index.json + requirements.json
# 1) 抽取（.doc 自动转换；DGS 环境改用 com_docx_extract.py）
python scripts/extract_docx.py "<docx路径>"
# 1b) 可选：Word 页码定位（Windows+Word；把 P段落号 映射为 Word 页码）
python scripts/page_map.py "_run/<文档名>"
# 2) 确定性检查（产出 det_*.json 机器证据）
python scripts/deterministic_checks.py "_run/<文档名>"
# 2b) 按组切片 + 直出 4 条结论（提速提质核心）
python scripts/slice_artifacts.py "_run/<文档名>"
python scripts/det_findings.py "_run/<文档名>"
# 3) 并行语义检查：Claude fan-out 4 个 subagent（各读 _slices/groupX.json，写 groupX.json）
# 4) 收集 + 合并 + 校验（det 覆盖 subagent，42 条全覆盖）
python scripts/collect_findings.py "_run/<文档名>"
# 4b) 结论润色（去除机器字段名黑话）
python scripts/polish_findings.py "_run/<文档名>"
# 5) 合成报告（默认 html；--format md 出 Markdown；--format both 两者都出）
python scripts/assemble_report.py "_run/<文档名>"            # 默认 HTML
python scripts/assemble_report.py "_run/<文档名>" --format=md  # 只要 Markdown
```

## 设计取舍

- **效率**：抽取与确定性检查为单进程秒级；语义检查 4 组并行（从原 7~8 组合并而来），每个 subagent 只读裁剪后的切片（<2 万 token，而非全文 8 万 token），墙钟≈最慢一组，整体提速 50~70%。
- **严谨**：数值/对照/编号类交由确定性脚本（零臆测），其中纯规则条目（std_001/003/004、gram_007）直接出结论不经 LLM；语义类强制引用原文定位与机器证据。
- **可信度**：确定性脚本对易误报场景（kW/台双单位、经验频率列、序号列、多合计行、区间值）仅给 `info`，是否成问题由语义结合网格判定。
- **规范依据**：流量有效数字按 SL/T 247《水文资料整编规范》判定（整数末尾零不算有效数字，2050 = 3 位）。
- **可移植**：`references/` 可替换为其他行业/版本的清单，索引自动重建。

## 依赖
Python 3 + `python-docx`、`lxml`。`.doc` 转换可选 LibreOffice 或 pywin32。
