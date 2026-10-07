---
name: design-report-review
description: 对水利工程设计报告（Word .docx/.doc）进行智能校审。按"检查要求.md"中强制性条文、常见设计问题、一致性、语法表述、文字规范性、设计标准、法律法规、表格逻辑关系共8大类42条逐项检查，输出与要求一一对应、引用完整要求原文的校审报告。默认输出仪表板式 HTML（顶部横幅+左侧目录+KPI统计卡+问题筛选+按类别表格化结论、行级配色），用户指定时输出 Markdown；两种格式均正确渲染上标/下标（如 m³/s、10⁸、km²、H₂O）。触发词：校审/审查/检查 报告、报告智能校审、docx校审、设计报告检查、校审报告。
---

# 报告智能校审 · 技能说明

本技能对一份水利工程设计报告（项目建议书/可研/初设等，Word 文档）执行**逐条、可追溯**的智能校审，覆盖 `references/检查要求.md` 的 **8 大类 42 条**检查要求，并按其"检查结果输出要求"产出校审报告：**默认输出 HTML**（用户明确要求 Markdown 时输出 Markdown）。每条要求都一一对应、不遗漏（即使全部符合也输出主要分析结论），并引用完整的各检查项要求原文。两种格式均**正确渲染上标/下标**（m3/s→m³/s、10的8次方→10⁸、km²、H2O→H₂O 等）。HTML 版为**仪表板式**：顶部横幅 + 左侧目录（滚动高亮）+ KPI 统计卡 + 问题筛选（全部/仅存在问题/仅不符合）+ 按类别表格化结论（每条一行、行级配色），便于速览与归档。

> 设计要点（提速版）：**确定性脚本**负责抽取与机器可验证的硬证据（数值、引用对照、表格核算、编号顺序），其中**纯规则条目直接出结论**（不再派 subagent）；**并行 subagent** 只跑语义判断（5 组 fan-out）。每个 subagent 只读**按组裁剪后的切片文件**（`_slices/` 下的组切片），而非全文。二者结合保证严谨、高质量与高效率。

---

## 0. 运行环境与路径

- 技能根目录（下称 `$SKILL`）= 本 SKILL.md 所在目录。
- 依赖：Python 3 + `python-docx`、`lxml`（已随环境安装）。Windows 下 `.doc` 转换可选依赖 LibreOffice 或 pywin32。
- 关键脚本：`$SKILL/scripts/` 下的 `extract_docx.py`、`deterministic_checks.py`、`slice_artifacts.py`、`det_findings.py`、`collect_findings.py`、`assemble_report.py`、`page_map.py`、`maintain.py`、`json_append.py`（subagent 增量写盘助手）。
- 参考/索引（首次运行自动生成）：`$SKILL/scripts/references_index.json`、`$SKILL/scripts/requirements.json`。
- 运行产物目录：`$SKILL/_run/<文档名去后缀>/`。

如索引/要求文件缺失，先运行：
```bash
cd "$SKILL" && python scripts/maintain.py --force   # 一键（等价于分别跑 build_reference_index.py + parse_requirements.py）
```

---

## 1. 流程总览（6 步）

```
[输入 docx]
   │  ① 抽取            extract_docx.py
   ▼
artifacts: fulltext.md / tables.json / captions.json / citations.json / numbers.json / meta.json / locators.json
   │  ①b 页码定位（可选，后台） page_map.py ←与③并行执行，⑤前汇合
   │  ② 确定性检查       deterministic_checks.py
   ▼
证据: det_table.json / det_citations.json / det_numbers.json / det_numbering.json
   │  ②b 按组切片        slice_artifacts.py        【新】
   │  ②c 直出4条结论     det_findings.py           【新】
   ▼
_slices/groupA1/A2/B/C/D.json（裁剪后小文件） + det_findings.json（std_001/003/004、gram_007）
   │  ③ 并行语义检查（5 组 subagent，每组只读对应 _slices/ 组切片；增量写盘；同时后台跑 ①b）
   ▼
各 subagent 增量写 groupX.json（骨架先行、每判完 1 条立即追加）
   │  ④ 收集+合并+校验   collect_findings.py       【新】
   ▼
findings.json（42 条全覆盖；det 覆盖 subagent；仍有缺失 → 退出码 1 硬门禁，占位 ≠ 完成）
   │  ④b 结论润色         polish_findings.py        【新】去除机器字段名
   ▼
findings.json（机器黑话已翻译为自然语言）
   │  ⑤ 合成报告         assemble_report.py
   ▼
[输出 <文档名>_校审报告.html（默认）/ <文档名>_校审报告.md（用户指定时）]
```

### ① 抽取
```bash
python "$SKILL/scripts/extract_docx.py" "<docx或doc路径>" ["<输出目录>"]
```
默认输出到 `$SKILL/_run/<文档名>/`。`.doc` 会自动尝试转换；转换失败则提示用户另存为 `.docx`。
抽取同时生成 `locators.json`：每段最近**章节标题**、每表**题录号(表X.X.X)** 映射。

### ①b 页码定位（可选，需 Windows + Word；后台与 ③ 并行）
```bash
python "$SKILL/scripts/page_map.py" "$SKILL/_run/<文档名>"
```
用 Word COM 计算每个顶层段落对应的 **Word 页码**，回填 `locators.json` 的 `page_by_para`（无 Word 时跳过，定位自动回退为章节标题§）。合成时 `P0154`→`第16页`、`T039`→`表3.7.2.2`。
**提速要点**：Word 分页较慢（约 1~2 分钟），而 `page_by_para` 只被 ⑤ 合成使用（②切片、③subagent 均不读它）——因此在 ②c 完成后**立即后台启动**（`run_in_background`），与 ③ 的 subagent 并行执行，⑤ 之前确认其完成即可，**墙钟零增加**。

### ② 确定性检查
```bash
python "$SKILL/scripts/deterministic_checks.py" "$SKILL/_run/<文档名>"
```
产出五份 `det_*.json` 证据并打印摘要（含 `det_suspects.json` 错字/叠字/标点疑点候选，由 `references/错字词典.json` 驱动，供诊断与词典召回评估，**不进切片**——实测候选核实会拖慢 C 组，2026-09-27 实验已回退）。**这些证据是数值类、对照类检查的权威依据，subagent 必须优先引用。**

### ②b 按组切片（提速核心）
```bash
python "$SKILL/scripts/slice_artifacts.py" "$SKILL/_run/<文档名>"
```
按 5 个分组裁剪产物，生成 `_slices/groupA1/A2/B/C/D.json`（大报告会自动限量：表格只给有问题的+采样、numbers 只给 moduli/flows、references_index 只给命中项、全文按疑点关键词裁剪）。**A1（强条 mp，11 条）与 A2（常见 ci，9 条）共用同一份 A 数据切片**——内容与拆分前 groupA 完全一致，仅提示词中要求清单不同；拆组只为把原 20 条最重组的负载减半，不丢任何证据。**subagent 只读切片，不再灌全文。**

每个切片内嵌 `"requirements"` 字段（本组各条的 id/name/content 原文，程序化取自 requirements.json）——**下发提示词直接引用切片内清单，禁止手抄要求清单**（2026-09-30 实战曾因手抄漏发 ci_runoff_002/ci_water_002/ci_water_003 共 3 条）。生成前做**分组覆盖自检**：5 组清单 ∪ det 直出 4 条 ≠ 全部要求即报错退出，分组定义与要求清单漂移当场拦截。

### ②c 直出 4 条结论（提质 + 提速核心）
```bash
python "$SKILL/scripts/det_findings.py" "$SKILL/_run/<文档名>"
```
把 4 条纯规则条目直接转成 findings 结构（`det_findings.json`），**这 4 条不再派 subagent**：

| id | 判定规则（来自 det_*.json） |
|---|---|
| std_001 | `std001_bigunit.mixed_form`：文字单位与10的次方混用→不符合（10的次方含"×10N"平文写法，直接扫全文，2026-10-01 修召回） |
| std_003 | `std003_flow.violations`：>3位有效数字或小数>3位→不符合 |
| std_004 | `std004_grade`：等别/航道等级/围岩及水质类别应为罗马专用字符、建筑物及堤防级别应为阿拉伯数字，另收拉丁字母 I/V/X 冒充罗马数字（如"III等"）。0 处→符合；1~3 处→部分符合；>3 处→不符合 |
| gram_007 | `det_numbering`：重复/断号→不符合；仅格式不统一→部分符合 |

> 理由：这些是数数/查表/加法，脚本比 LLM 更准（实测 std_003 脚本抓 23 处 vs LLM 15 处）。SKILL 设计原则"宁可漏报也不误报"，det 给的 fail 是高置信硬伤。

### ③ 并行语义检查（5 组 fan-out，效率核心）
**在一条消息内并发启动 5 个 subagent**（Agent 工具一次发 5 个调用）。每个 subagent 只读**一个切片文件** `_slices/groupX.json`（X∈A1,A2,B,C,D），互不阻塞。

| 组 | 检查项 | 条数 | 读切片 | 附 det 证据（已在切片内） |
|---|---|---|---|---|
| **A1** 强条 | mp_001–011 | 11 | groupA1.json | det_std003_brief |
| **A2** 常见 | ci_flood/runoff/stage/water/drainage | 9 | groupA2.json | det_std003_brief |
| **B** 一致性+表格 | cons_001–003、tbl_001–003 | 6 | groupB.json | det_table_brief(fail/warn/info) |
| **C** 语法+文字 | gram_001–006/008、std_002 | 8 | groupC.json | det_numbering、det_std004_grade |
| **D** 标准+法规 | ds_001–002、lr_001–002 | 4 | groupD.json | det_citations_brief(summary+可疑清单) |

> A1/A2 的数据切片与拆分前 groupA 完全一致（fulltext 关键词切片 + flows/moduli + std003 摘要），拆组不减证据、只减每组条数。

> **不派 subagent 的 4 条**：std_001、std_003、std_004、gram_007（由 ②c 直出）。
> **半直出条目**（tbl_001/ds_001/ds_002/lr_001/lr_002）：det 硬证据已在切片内，subagent 直接引用 det 结论，不重复核算。
> **大报告**：groupC 切片会提示"用 Grep 分段读 fulltext.md"检查标点/错字，勿一次性读全文。

**验收标准（completed ≠ 成功）**：一个组"完成"的唯一定义是——`group<X>.json` 存在、`python json_append.py <文件> verify` 通过、且 results 条数 ≥ 切片 requirements 字段的长度。运行界面显示 completed **不构成**完成依据：turn 结束但文件缺失或缺条目一律视为失败（2026-09-30 实战 B 组显示 completed 却没写文件）。

**失败恢复三梯度（按序尝试，成本递增；2026-09-30 实战 6/6 靠梯度 1 收齐）**：
1. **resume 收尾模式**：对失败/超时的运行直接 resume（恢复自带已积累的分析上下文，成本最低），指令改为收尾模式——禁止再做新的系统性检索/阅读；已分析未写盘的条目立即用 json_append.py 补写；未分析完的条目最多 ≤3 处定点读取后给简短判定；verify 通过后立即结束。
2. **主会话定点补判**：仍缺的条目由主会话在 fulltext.md 定点 Grep 检索关键词后，用 json_append.py 写入 `group<X>b.json`（collect 自动收集 run_dir 下任意 `group*.json`，拆分命名无需适配）。
3. **重发更小任务**：前两梯度无效才重新派发，且必须进一步拆小（≤6 条/组，见 §5 拆分预案）。

### ④ 收集+合并+校验（覆盖硬门禁）
```bash
python "$SKILL/scripts/collect_findings.py" "$SKILL/_run/<文档名>"
```
收集 5 组 subagent 的 groupX.json + det_findings.json，校验字段、**det 覆盖同 id 的 subagent 结果**、对照 requirements.json 确保 42 条全覆盖。**全覆盖是硬门禁**：仍有缺失时 findings.json 照写（缺失条目按"未检出"占位、`_source=missing`），但脚本**打印 FAIL 并以退出码 1 结束**——占位 ≠ 完成，须按 ③ 的失败恢复三梯度补齐后重跑 collect，不得带占位进入 ⑤；确属无法补判时加 `--allow-missing` 放行（最终报告须向用户说明）。注意：退出码 1 是门禁信号不是脚本崩溃，先看输出再决定补判方式。

### ④b 结论润色（必跑）
```bash
python "$SKILL/scripts/polish_findings.py" "$SKILL/_run/<文档名>"
```
subagent 有时会把切片里的 det 证据字段名（`det_citations_brief.summary`、`consistency_issues=0`、`grid_omitted`、`n_fail` 等）直接照抄进给用户看的结论。本脚本把这些机器黑话翻译成自然语言（"机器核算 0 项不合格""未发现前后写法不一致"等），并清理替换后产生的不通顺。**必须在 collect 之后、assemble 之前运行**，原地更新 `findings.json`。

### ⑤ 合成报告
```bash
python "$SKILL/scripts/assemble_report.py" "$SKILL/_run/<文档名>" [--format md|html|both]
```
默认 HTML；用户明确要求 Markdown 时加 `--format md`。`assemble_report.py` 内置 det 双保险（det_findings 覆盖优先），并给 det 来源条目标注"（机器判定）"。每条检查项的问题清单默认最多列 30 条，可用 `--max-issues=N` 调节；超限时列表末尾会明示"共 X 条，其余 Y 条未显示"。

---

## 2. 产物与证据字段速查

- **fulltext.md**：全文，每段前缀 `P0123`（段落序号）/`[H3]`（标题层级）；表格以 `【表012】(r×c) 题录（详见 tables.json#12）` 标记。→ 语义判断与定位引用（subagent 一般不直接读，走切片）。
- **tables.json**：`[{id,pid,caption,nrows,ncols,header,grid,section}]`，`grid` 为二维文本（合并单元格已重建）。→ 表格检查（切片已含问题表 grid）。
- **captions.json**：`[{type:图|表,num,text,is_def,para_idx,section}]`。→ 题录/编号检查。
- **citations.json**：`[{kind,std|law|law_code|std_name_only,raw,norm_code,name,norm_name,para_idx,context}]`。→ 引用检查的原始数据。
- **numbers.json**：`[{key:flows|moduli|areas_km2|areas_mu|bigunits|pow10,value,raw,para_idx,context}]`。
- **det_table.json**：每表 `{id,caption,n_checks,n_fail,n_warn,checks:[{type,stated,computed,status,note}]}`。`status` 含 `pass/fail/warn/info`；`info`（行平均、极值、多合计行/列）为机器核算供语义复核，**不直接判错**。
- **det_citations.json**：`{citations:[{kind,raw,norm_code,name,validity,note,...}],consistency:[...],summary}`。`validity`：`valid/not_in_list/year_mismatch/valid_name`。
- **det_numbers.json**：`std003_flow`(流量有效数字违规)、`std001_bigunit`(大数单位形式)、`std004_grade`(等别/级别罗马-阿拉伯)。
- **det_numbering.json**：`{图/表:{count,formats,format_inconsistent,issues}}`。
- **det_suspects.json**：`{summary, candidates:[{rule,loc,span,hint}]}`，错字词典+叠字/标点/配对/长句五类规则产的疑点候选（**只产候选不产结论**）。安徽两湖实测对 C 组金标准召回 28%，未接入切片（实验记录见 tests/README.md）。
- **references/错字词典.json**：错字词典 76 条四级分档（挖掘自 414 份平台历史校审报告 + 20 个运行目录 + 行业补录），驱动 det_suspects 词典规则；增量维护用工作区 `词典/harvest_typos.py` 重跑。
- **_slices/groupA..D.json**：按组裁剪的切片（见 ②b），内嵌本组 `requirements` 要求清单。**subagent 的唯一输入**。
- **det_findings.json**：4 条直出结论（std_001/003/004、gram_007），格式同 findings.json 子集。
- **references_index.json**：标准/法规按 `norm_code`、`norm_name`、`prefix_num` 检索（切片已含命中项，subagent 不直接读全索引）。
- **requirements.json**：`[{category,subcategory,id,name,content}]`（42 条）。

---

## 3. Subagent 提示模板与输出 schema

给每个 subagent 的提示（按组填充 `<...>`，**只读一个切片文件**）：

```
你是水利工程设计报告校审专家。请对以下检查要求【逐条】给出结论。
只使用我提供的切片数据 与（必要时 Grep 读取的）文档原文，禁止臆测；每条结论必须给出文档中的定位(段落号 PXXXX / 表号 / 题录)与原文片段作为依据。

【本组检查要求】（必须全部覆盖，逐条响应）：
以切片文件内的 "requirements" 字段为准（本组各条的 id/name/content 原文，共 N 条）——
以下发的切片为准、不得增删、不得凭记忆另写清单（2026-09-30 实战曾因手抄清单漏发 3 条）。
仅当切片为不含 requirements 字段的旧版时，才从 requirements.json 粘贴本组各条原文。

【唯一输入】读取这一个文件（含本组所需的裁剪后全文/表格/数值/det 证据）：
$SKILL/_run/<文档名>/_slices/group<X>.json

【判定口径】
- 符合：报告完整满足该要求。
- 部分符合：基本满足但存在不足/疏漏。
- 不符合：存在明确问题或缺失。
- 不适用：该要求不适用于本文档（须说明原因）。
- 即便全部符合，也要写出主要分析结论与支撑定位。
- 涉及数值/标准/法规时，切片内已附 det 证据摘要，直接引用，不重复核算。
- **不适用判定纪律**：判"不适用"前必须先在切片的全文/关键词切片中检索该项关键词（如 水面线、壅水、糙率、洪痕、考证、率定、插补、比拟法、分期、高程系统）；凡文档中存在与该要求对应的分析或计算内容——即使只是汇编性描述、背景引用、施工期或局部河段内容——一律视为适用并按符合程度判定；只有全文完全无相关内容才可判"不适用"，且结论中须列出已检索而未命中的关键词。
- **判定尺度**：要求中的每个动作词（说明/分析/论证/率定/对比/校核/考证）在正文中缺失或只完成一部分，该项即至少判"部分符合"并逐处列出缺口；不得因正文提及相关主题、资料名或规范名而判"符合"。
（实测教训 2026-09-27：无此两道防线时，拆组后单组曾把文档中已有的水面线推算、分期施工期洪水误判"不适用"，漏报约 9 处问题；加入防线后 issues 由 18→40，反超旧架构基线的 26。）

【具体问题清单——关键要求，避免空泛】
对"部分符合/不符合"的检查项，**必须**逐处列出具体问题（不要只写"需逐处订正"这类空话），每条给：
- loc：段落号 PXXXX（合成时会自动转成 Word 页码"第N页"）
- original：从原文**精确照抄**的问题片段（含错字/原样）
- type：问题类型（2-6字，如"叠字/错字/标点/量纲/前后矛盾/编号倒置"）
- suggestion：**逐字修改建议**——直接给出修改后的文字 + 简述改法
尽量穷尽明显问题（每项可多条，宁缺毋滥、不编造）。

【⚠ 结论措辞要求】conclusion 是给报告作者看的，**禁止**出现切片文件的内部字段名或自称。
不要写 "det_citations_brief.summary 显示 consistency_issues=0""grid_omitted:true"
"n_fail=0""空数组" 这类机器黑话，也不要写 "本组切片""本组""所提供切片" 这类暴露
内部实现的措辞——一律用"本文档/报告"自称。要把机器证据翻译成自然语言，如"机器核对
未发现前后不一致""表格未提供单元格数据""0 项不合格"。可引用具体数值和定位，但不要
暴露字段路径或切片概念。语法(gram_*)、一致性(cons_*)等尤其需要这种可操作清单。

【大报告提示】若切片含 _fulltext_note（全文过大），gram_001–006（标点/错字/通顺）请用 Grep 在
<run_dir>/fulltext.md 按段落号区间分段（如 P0000–P0500）逐段读取检查，每段报具体 loc+original+suggestion。

【输出——增量写盘纪律（严格遵守；运行可能中途被截断/超时，2026-09-30 实测）】
输出文件：$SKILL/_run/<文档名>/group<X>.json，逐条 result 结构：
  {"id":"<编号>","name":"<名称>","verdict":"符合|部分符合|不符合|不适用",
   "conclusion":"<结论，含主要分析>","evidence":[{"loc":"<定位>","quote":"<原文片段>"}],
   "issues":[{"loc":"P0238","original":"…","type":"叠字","suggestion":"…"}]}
1. 判定开始前先创建骨架（文件已存在会拒绝覆盖，保护可抢救成果）：
   python "$SKILL/scripts/json_append.py" "$SKILL/_run/<文档名>/group<X>.json" init "<组名>"
2. **每判定完 1 条，立即追加写入**（禁止攒到最后一次性写；禁止用 Edit 工具追加——
   多行中文锚点必失配）。该条 JSON 用 heredoc 从 stdin 传入：
   python "$SKILL/scripts/json_append.py" "$SKILL/_run/<文档名>/group<X>.json" <<'EOF'
   {"id":"…","name":"…","verdict":"…","conclusion":"…","evidence":[…],"issues":[…]}
   EOF
   （运行中途死掉也不亏：已追加的条目都已落盘，恢复运行只补缺条。）
3. 全部判完自检后立即结束：
   python "$SKILL/scripts/json_append.py" "$SKILL/_run/<文档名>/group<X>.json" verify <本组条数>
4. 除 group<X>.json 外**禁止创建/写入任何其他文件**（含临时文件与临时目录——首败根因
   即子代理往不存在的 G:\tmp 写临时文件）；中间笔记一律并入 group<X>.json 的 conclusion。
```

> 5 个 subagent 各写一个 `groupA1/A2/B/C/D.json`（category 取本组对应的检查类别名：A1="强制性条文检查"、A2="常见设计问题检查"；B 组的"一致性检查"+"表格逻辑关系专项检查"可拆成两个 results 数组或分两个文件，collect_findings 都能识别）。按 §5 拆分预案派发的小组写 `groupA1a/A1b/B1/B2/C1/C2.json` 等任意 `group*.json`，collect 自动收集。**收齐的判据是 ③ 的验收标准（文件存在 + verify 通过 + 条数齐），全部达标才进入 ④。**

### ⚠ 流程纪律（禁止现场拼脚本）

- 中间产物一律走 `slice_artifacts.py` / `collect_findings.py`。
- **不得**在 `_run/` 下创建 `build_findings.py`、`_merge.py`、`merge_findings.py`、`extra_*.json` 等临时脚本或临时合并文件——这些是历史绕路产物，已由 collect_findings 取代。
- subagent 的输出必须是符合 schema 的 `groupX.json`，不得输出其他命名。

---

## 4. 报告输出格式（严格遵守"检查结果输出要求"）

- 与要求**一一对应、不遗漏**：按 `requirements.json` 顺序，42 条全部出现；collect_findings 的覆盖硬门禁保证（缺条=FAIL 退出码 1，占位条目不得带入报告）。
- **引用完整要求原文**：每条在结论前以引用块给出该条要求全文（取自 requirements.json 的 content；强制性条文条目含规范条款原文）。
- 即便符合，也写"主要分析结论"。
- **默认输出 HTML**，保存到**输入文档同目录**：`<文档名>_校审报告.html`（**仪表板式**：顶部横幅 + 左侧目录(滚动高亮) + KPI 统计卡(总数/符合/存在问题/符合率) + 问题筛选(全部/仅存在问题/仅不符合) + 按类别表格化结论(列：检查项｜要求｜检查结果｜说明/建议，每条一行、行级配色 绿/橙/红/灰、有问题内容红色粗体)，说明/建议列只给结论与建议（不单列定位依据引用），浏览器直接打开/归档、支持打印）。
- **用户明确要求 Markdown 时**才输出 `<文档名>_校审报告.md`。
- 由 `assemble_report.py` 控制：默认 `--format html`；`--format md` 出 Markdown；`--format both` 两者都出。
- **两种格式均正确渲染上标/下标**：`m3/s`→m³/s、`10的8次方`→10⁸、`km2/km²`→km²、`H2O`→H₂O、`SO4`→SO₄ 等。HTML 用 `<sup>`/`<sub>` 标签，Markdown 用 GFM 内联 `<sup>`/`<sub>` 标签（在 GitHub/VSCode/Typora/Obsidian 等渲染器中正确显示）。

模板：

```markdown
# 《<文档名>》智能校审报告

- 文件：<文档名>　段落：<n>　表格：<n>　标题：<n>
- 校审依据：检查要求.md（8 大类 42 条）；标准清单：水利标准.md；法规清单：法律法规.md
- 生成方式：确定性脚本（抽取+机器证据+4条直出）+ 5组并行语义检查

## 一、强制性条文检查
### mp_001 设计洪水计算过程检查
> **要求**：《水利水电工程设计洪水计算规范》1.0.9：……（完整原文）
**结论**：符合 / 部分符合 / 不符合
**依据**：P0123「……原文片段……」；……
（…… mp_002 … mp_011 ……）

## 二、常见设计问题检查
### ci_drainage_001 排涝模数合理性分析
> **要求**：对设计排涝流量要进行合理性分析：5年一遇农田涝区0.50～0.65m³/s·km²……
**结论**：部分符合
**依据**：numbers.json 记录排涝模数54处；其中 P0586「排涝模数0.33m³/s/km2，不足5年一遇」低于5年一遇农田下限0.50；P…「3.2、4.4」显著超城镇20年一遇上限1.40，需核实……

## 三、一致性检查 / 四、语法表述检查 / 五、文字规范性检查 / 六、设计标准检查 / 七、法律法规检查 / 八、表格逻辑关系专项检查
（同样格式，逐条展开）

---
## 校审结果汇总
| 类别 | 条数 | 符合 | 部分符合 | 不符合 | 不适用 |
| ... | ... | ... | ... | ... | ... |
```

---

## 5. 效率与降级

- **首选 5 组并行**：在**一条消息内**同时发起 5 个 subagent（A1/A2/B/C/D），墙钟≈最慢一组；原 A 组（20 条最重）已拆为 A1（11 条）+A2（9 条），共用同一数据切片，负载均衡且证据不缩水。
- **组规模与拆分预案（≤6 条/组最稳）**：实测（2026-09-30，glm-5.3-flash + 214 页报告）11 条的 A1、8 条的 C 在单次运行内必死于输出超长/30 分钟超时；6 条左右的小组配合增量写盘全部可收齐。**模型较弱（flash 级）或文档很大（>1500 段或切片 >150KB）时，把标准 5 组拆成 8 组**：A1→A1a（mp_001~006）+A1b（mp_007~011）、B→B1（cons_001~003）+B2（tbl_001~003）、C→C1（gram_001~004）+C2（gram_005/006/008/std_002）。拆分后各小组写 `groupA1a.json` 等任意 `group*.json`（collect 自动收集），**切片不用重切——同组拆分共用原切片**（A1a/A1b 都读 groupA1.json，B1/B2 读 groupB.json，C1/C2 读 groupC.json），要求清单以切片 requirements 中各自负责的条目为准。
- **page_map 后台并行**：①b（Word 分页约 1~2 分钟）只被 ⑤ 使用，与 ③ 同时后台执行，不占墙钟。
- **切片已限量**：subagent 无需自行判断读什么。
- **大报告**：groupC 切片会提示用 Grep 分段读 fulltext.md 查标点/错字；groupB 表格只给有问题的+采样。
- **降级**：若无法并发，按 A1→A2→B→C→D 顺序逐组串行，结果不变。
- **证据优先**：凡 det_*.json 已给出的结论（std_003 流量有效数字、std_001 大数单位、std_004 等别级别、gram_007 编号——由 det_findings 直出；tbl_001 表格 fail、ds 标准有效性——切片内附摘要），subagent 直接引用、不重复人工核算，避免误判。

---

## 6. 回归测试（发布纪律）

`$SKILL/tests/` 下有金标准测试集与门禁脚本，**任何对 SKILL.md、subagent 提示词、scripts 的改动，发布前必须过门禁**：

- 金标准：`tests/golden/shanxu`（小文档，109 处问题）、`tests/golden/anhui`（超大文档，236 处问题），各含冻结的 `artifacts/`（抽取产物+det 证据+切片）与 `findings_golden.json`。**只读**；更新属刻意行为，须重校准并在 golden_meta.json 记录原因。
- findings 门禁：`python tests/regression_check.py findings <run_dir> --golden tests/golden/<名>`，五项检查——42 条覆盖且未检出=0；verdict 翻转分类（**"部分符合/不符合→符合/不适用"的放宽方向 = FAIL**，漏报信号）；各条 issue 数量下限（金标准 ≥5 条的项须 ≥80%）；反幻觉抽查（issue 的原文片段与 loc 段落号必须能在产物中找到，det 直出条目除外）；机器黑话=0。
- artifacts 门禁：`python tests/regression_check.py artifacts <run_dir> --golden ... [--allow groupC,det_suspects]`，确定性产物与切片逐项比对，预期变更用 --allow 豁免。
- 纪律：小改动跑 shanxu；涉及切片/分组/提示词的改动跑 anhui 全流程。**FAIL 即回滚或修复**，不得为放行而放宽门槛（门槛调整本身须记录并重校准）。

---

## 7. 边界与注意

- `.doc`（旧二进制）：脚本自动尝试 LibreOffice/Word 转换；失败则要求用户另存 `.docx`。
- 表格合并单元格已重建为二维网格；`det_table` 对多合计行/列、kW/台双单位、区间值等只给 `info`，**是否成问题由 G 组 subagent 结合 grid 判定**。
- 引用清单为"现行有效版本"参考；`not_in_list` 可能是地方/行业外标准或内部技术文件，subagent 应区分"编号错误/已废止"与"非本清单范围但合理"。
- 强制性条文（mp_*）判"落实"须以正文是否对该环节做了分析/论证为准，并给出定位；不得仅因提及规范名即判符合。
