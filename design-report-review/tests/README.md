# 回归测试集与门禁（design-report-review）

## 组成

```
tests/
  regression_check.py   门禁脚本（findings / artifacts 两种模式）
  golden/
    shanxu/   山许水库项目建议书（小文档代表，109 处问题，2026-09-27 冻结）
    anhui/    安徽两湖涝区项建第2、3、5章（超大文档，236 处问题，2026-09-27 冻结）
      artifacts/          抽取产物 + det_*.json + _slices/（冻结切片）
      findings_golden.json 金标准结论
      golden_meta.json     冻结记录（版本/来源/原因）
```

## 用法

```bash
# findings 门禁（改了提示词/分组/collect/polish 后必跑）
python tests/regression_check.py findings <run_dir> --golden tests/golden/anhui

# artifacts 比对（改了 extract/det/slice 后必跑；预期变更用 --allow 豁免）
python tests/regression_check.py artifacts <run_dir> --golden tests/golden/anhui --allow groupC
```

退出码 0=PASS，1=FAIL。**FAIL 即回滚或修复；不得为放行而放宽门槛。**

## 门禁五项检查（findings 模式）

1. 42 条全覆盖、未检出=0
2. verdict 翻转分类：放宽（部分符合/不符合 → 符合/不适用）= FAIL（漏报信号）；收紧/新覆盖 = WARN
3. issue 数量下限：金标准 ≥5 条的项须 ≥80%；<5 条的项容差 n-2
4. 反幻觉抽查（确定性采样 ≤40 条）：issue.original 的片段必须在 fulltext/tables/captions 中找到
   （空白不敏感、上下标归一、引文片段抽取、数据型对照验数字），loc 段落号必须存在
5. 机器黑话残留 = 0

## 校准记录（2026-09-27）

- 双金标准自测 PASS；artifacts 自一致 PASS。
- 负样本校准：以 7 月旧结论冒充新结果 → 正确 FAIL（3 条放宽翻转 + 21 项 issue 不足 + 总量 136<80%×236）。
- 门禁自身抓到的真问题：金标准 mp_002 的 P0446 段落号写岔（实为 P0445，subagent off-by-one，
  已修正冻结）；山许金标准 5 条机器黑话残留（polish 规则盲区，已补规则并重清洗）。

## 杠杆 3 实验记录（det 疑点候选攻 C 组，2026-09-27，结论：两案皆否）

1. **预扫替代通读**：召回实测仅 46/163=28%（语病/歧义/逻辑/常识类为 0，必须语义阅读）→ 否决。
2. **候选核实+通读并存**：C 组实测 1711s / 5.8M tokens（基线 893s / 1.3M），候选核实成为
   额外一遍工作；issues 226>163 且过门禁（质量面合格），但效率面回退 → **切片集成已回退**，
   det_suspects.json 生成器保留作诊断工具（叠字类召回 5/6 值得后续做 det 直出实验）。
3. 附带产出：v3 的 226 条 C 组问题过门禁后已用于升级安徽两湖交付报告（总 299 条问题）。

## 新增金标准流程

选一个有代表性的新报告 → 跑完整管线（含 subagent）→ 人工抽查 findings 质量 →
`cp -r <run_dir>/artifacts tests/golden/<名>/ && cp <run_dir>/findings.json tests/golden/<名>/findings_golden.json`
→ 写 golden_meta.json → 用门禁自测校准。

## v1.4.0 可靠性改造（2026-10-01，源自 9-30 赣东北可研实战教训）

背景：2026-09-30 赣东北可研（1094 段 120 表，glm-5.3-flash）5 组整发 3 组死亡
（临时文件 ENOENT / 假 completed / 30min 超时），靠现场"拆小组+增量写盘+resume 收尾"
抢救收齐。本次把现场手段固化为 skill 默认行为：

1. **json_append.py**：subagent 增量写盘助手（init 骨架 / heredoc 逐条追加 /
   verify 自检；重复 id 拒绝、损坏文件拒绝写入）。全部分支自测通过。
2. **slice_artifacts.py**：切片内嵌本组 requirements（程序化取自 requirements.json，
   杜绝手抄清单——实战曾漏发 3 条 ci_*）；分组覆盖自检不等即退出。
3. **collect_findings.py**：覆盖升级为硬门禁——缺条仍写 findings.json 但退出码 1，
   `--allow-missing` 才放行。
4. 全部入口脚本加 GBK 控制台自愈（stdout/stderr errors=replace）。
5. SKILL.md：③ 加验收标准（completed≠成功）+ 失败恢复三梯度（resume 收尾→定点
   补判→重发更小）；§5 加 ≤6 条/组拆分预案（A1a/A1b、B1/B2、C1/C2 共用原切片）。

**金标准切片重冻结（刻意行为）**：两套 golden 的 _slices 用 v1.4 代码重生成
（仅新增 requirements 字段），golden_meta.refreeze_log 已记录。shanxu 顺带修复一处
陈旧不一致：旧冻结切片的 det_std003_brief 仍是 sig-figs 修正前的 10 违规版本，
与同目录 det_numbers.json（0 违控）矛盾。

**验证**：金标准反拆回灌测试——38 条 subagent 结论经新 collect+polish 逐字节一致；
artifacts 门禁 PASS。**发现一处先于本次改动的 det 校准欠账（未修，另行处理）**：
金标准 4 条 det 直出条目是 LLM 时代陈旧结论（std_003 结论还引用旧产物名"数值检查.json"），
与现行 det 行为冲突——std_001（金标 不符合/10 issues vs 现行 det 符合/0）根因是
**抽取召回缺口**：山许文档实有 25 处"×104"式 10 的次方误写，numbers.json pow10=0
全未捕获（std_001 混用判定因此漏报）；std_004 现行 det 只扫显式"等别/级别"表述
（全文仅 1 处命中，"N级"类表述不在扫描范围）；gram_007 现行 det 发现表 3.5 编号
间断 2→5→8 而金标准误判"符合"。任何真实重跑都会在 std_001/std_004 触发放宽 FAIL
——这正是门禁该抓的。待办：先修 pow10/"N级"抽取召回，再重生成 det 证据并重校准
金标准 4 条 det 条目。

## det 召回修复 + 金标准 det 条目重校准（2026-10-01 第二批）

承接上节"det 校准欠账"，完成三件事：

1. **std_001 召回修复**：pow10 计数从 extract 分类的 numbers.json 改为 det 直接扫
   fulltext.md（"×10N"平文写法 + "10的N次方"字面写法，指数限 3~12 防"50×100"类
   量纲乘积误收）。山许实测 28 处 ×10⁴ 全部入证（P0220/0285/0330/0528/0529/0537/
   0545/0553/0557/0565），与旧金标准 10 处 issue 的段落集合完全吻合——旧金标准结论
   "检测器仅匹配'10的N次方'字面文本，未识别'×104'写法"的指认得到证实并修复。
   extract 的 numbers.json 不动（金标准无法重生成该文件，且 std001 已不依赖它）。
2. **std_004 扫描扩充**（高精度规则）：拉丁字母 [IVX]{2,} 冒充罗马数字（"III等"）、
   罗马数字+级（仅建筑物/堤防语境，阶地/航道用罗马属正确不收）、围岩/地表水/水质
   阿拉伯+类、阿拉伯+级（航道/船闸）。判定尺度：0 处符合 / 1~3 处部分符合 /
   >3 处不符合（个别笔误与成片误用区分）。山许仅命中 1 处（P0285 "III等"，与旧
   金标准指认一致），安徽零命中零误报。
3. **金标准 det 条目重校准**：两套金标准的 4 条 det 直出条目替换为现行 det_findings
   输出（结论文本机器可复现，不再保留 LLM 时代手写结论——旧 std_003 结论还引用
   "数值检查.json" 旧产物名）。山许 issues_total 109→126（std_001 10→28、std_004
   2→1）；唯一 verdict 翻转 gram_007 符合→不符合：核实表 3.5 题录序列 1,2,5,8,9，
   断号 2→5→8 真实存在，旧金标准只讨论了格式混用属漏判。安徽 236 不变、零翻转。
   det_numbers.json 与 _slices 已重冻结（refreeze_log 第三条记录）。

**门禁复核**：双金标准反拆回灌全流程（det_findings→collect→polish→两模式门禁）
全部 PASS（rc=0 ×4），"放宽方向翻转"0 条——det 欠账清零，真实重跑的 findings
门禁不再会被 std_001/std_004 拦截。
