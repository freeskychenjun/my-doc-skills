# my-doc-skills

处理文档的 AI agent skill 合集（Claude Code / DSH 等通用 `.agents/skills` 结构）。

## Skills

| Skill | 功能 |
|-------|------|
| [docx-paragraph-polish](docx-paragraph-polish/) | 校正 Word 文档正文段落：逐段润色/逻辑重构，以 Word 修订（track changes）模式写回，纯 OpenXML 实现，不依赖本机安装 Word/WPS |
| [design-report-review](design-report-review/) | 水利工程设计报告智能校审：8 大类 42 条逐项检查，输出仪表板式 HTML 或 Markdown 校审报告 |
| [design-report-format](design-report-format/) | 一键排版中文设计/工程报告（Word .doc/.docx）：页面、标题、正文、图表名、表格边框 |
| [wps-word-cooperate](wps-word-cooperate/) | **在 WPS 里和 AI 一起改 Word**：打开着的文档直接改，文字改动是可见修订，排版/表格/题注/目录/批注/版式核对一并做；带"拍快照→人裁决→读 diff→出交接单"的协作闭环。仅 Windows + WPS（走 COM，27 条命令） |

## 安装

```bash
git clone https://github.com/freeskychenjun/my-doc-skills.git
# 把需要的 skill 目录复制到（或 junction 到）你的技能目录：
#   Claude Code / DSH:  ~/.agents/skills/<skill-name>
```

各 skill 的环境要求与用法见各自目录下的 `SKILL.md`。

## 仓库约定

- 每个 skill 一个顶层目录，目录名即 skill 名。
- `design-report-review`、`design-report-format` 以 git subtree 迁入（保留完整历史），今后在本仓库统一维护。
- `wps-word-cooperate` 需要 **Windows + WPS Office**（走 COM 自动化），其余三个是纯文件处理、跨平台。
- 每个 skill 自带 `README.md`（给人看）、`SKILL.md`（给 AI 看）；`wps-word-cooperate` 另附单文件 `guide.html` 使用说明。
