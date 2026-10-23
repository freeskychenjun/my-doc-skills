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

## 致谢

- [copylee711/dsh-office](https://github.com/copylee711/dsh-office)（MIT）：`wps-word-cooperate` 的**做法**取自它——用 COM 让 Office/WPS 自己动手改文档、用页面图元（EMF）而不是截图来核对版式；`wps-word-cooperate/reference/wps-quirks.md` 里那份 WPS 与 Word 的差异清单，有相当一部分是通过观察它的行为记录归纳的。**代码没有拷贝它的**：helper 是按本项目的需求从零写的（C# 5、7 个 partial class 文件），实现路径也做了不同的取舍（例如它用常驻进程 + JSON 管道，本 skill 用"一命令一进程 + UTF-8 文件往返"，换来零安装与易分发）。
- [LibreOffice](https://www.libreoffice.org/) 与 Python 生态（`python-docx`、`PyMuPDF`、`Pillow`）在测试与文档生成环节提供了工具链支持；`design-report-format` / `design-report-review` 两个 skill 的部分排版与渲染思路来自这些工具的行为对照。
- 第 `docx-paragraph-polish` 的规则库与校审条目，来自本院历年报告的校审实践，不依赖任何第三方库。

## 许可证

本仓库采用 **MIT License**，见 [LICENSE](LICENSE)。各 skill 目录内若另有 `LICENSE`（如 `design-report-review/`），以目录内那份为准。
