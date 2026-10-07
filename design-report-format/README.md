# 设计报告一键排版 Skill

一个 [Claude Code](https://claude.ai/code) 技能（skill），通过 Word COM 自动化对 Word 设计/工程报告（`.doc`/`.docx`）一键排版：页面设置、一至四级标题、正文、图名/表名、表格，全部按规范统一。

排版引擎源自 Word 插件「CJTools 一键排版」的 `WordFormatter`，本仓库打包为**独立运行**的命令行程序，自带全部依赖，**无需安装任何插件**。

---

## 功能特性

- **页面设置** — A4，页边距 上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm。
- **标题** — 自动识别一至四级标题（`第X章`、`1.1`、`1.1.1.1` 等，含 Word 自动编号），应用字体字号并在导航窗格建立正确层级。
- **正文** — 宋体小四、两端对齐、首行缩进 2 字符、固定行距 25 磅。
- **图名 / 表名** — 识别「图X-X」「表X-X」短段落，黑体五号居中。
- **表格** — 外边框（默认 **1.5 磅**）、自适应窗口宽度、**字体默认跟随正文**（正文为宋体时西文用 Times New Roman，**字号不变**）、**首行标题自动外提**到表外、可选清理多余空格。
- **跳过页面** — 封面、扉页等前 N 页可不排版。
- **长文档进度** — 排版真实报告时，AI 默认**中途分段汇报进度**（处理到第几段、已用时）。

> 标题/图名表名的识别规则为引擎内置、不可配置：超过 60 字不识别为标题，超过 40 字的「图/表」开头段落不识别为图名表名。

---

## 环境要求

| 项 | 要求 |
|---|---|
| 操作系统 | Windows |
| Word | 桌面版 Microsoft Word **2010 及以上**（2013/2016/2019/2021/Microsoft 365） |
| 文档格式 | `.docx` 或 `.doc` |

不支持 WPS Office 和 Word 网页版。

---

## 用法

### 方式一：作为 Claude Code 技能（推荐）

直接用自然语言告诉 AI，例如：

- `排版 D:\报告\七星水库放空评估报告.docx`
- `排版这份报告，跳过前 4 页`
- `表格内容字号统一成五号`

AI 会自动调用本技能、分段汇报进度、完成后给出统计，并把结果**另存为新文件**（文件名加 `-排版` 后缀），原稿不动。

### 方式二：直接命令行

```bash
scripts/FormatDocx.exe <输入.docx> [--rules 规则.json] [--output 输出.docx]
```

- 不带 `--output`：原地保存（覆盖原文件）。
- 带 `--output`：另存为新文件（推荐）。
- 不带 `--rules`：依次尝试 `--rules` 指定文件 → `%AppData%\CJTools\rules.json` → 内置默认规则。本仓库的默认规则在 `references/default-rule.json`，使用时显式传入 `--rules references/default-rule.json` 即可。

**输出契约**（供集成方）：

- **stdout**：成功时输出**一行 JSON** 统计，例如
  ```json
  {"ok":true,"file":"...","rule":"水利工程报告","paragraphs":2899,"headings":30,"body":288,"images":13,"figCaptions":12,"tblCaptions":14,"extractedTableTitles":7,"skipped":2542,"elapsedMs":22316}
  ```
- **stderr**：进度日志，前缀 `[FormatDocx]`，每处理 50 段一行（含阶段、当前/总数、已用时）。
- **退出码**：`0` 成功 · `1` 排版出错 · `2` 参数/规则文件错误 · `3` 文档被占用、只读或受保护视图。

---

## 默认排版规范

| 段落类型 | 中文字体 | 西文字体 | 字号 | 其他 |
|---|---|---|---|---|
| 一级标题 | 黑体 | Times New Roman | 三号 16pt | 居中，段前段后 24 磅 |
| 二级标题 | 楷体 | Times New Roman | 四号 14pt 加粗 | 左对齐，段前段后 12 磅 |
| 三级标题 | 宋体 | Times New Roman | 小四 12pt | 左对齐 |
| 四级标题 | 宋体 | Times New Roman | 小四 12pt | 左对齐 |
| 正文 | 宋体 | Times New Roman | 小四 12pt | 两端对齐，首行缩进 2 字符，固定行距 25 磅 |
| 图名 / 表名 | 黑体 | Times New Roman | 五号 10.5pt | 居中，固定行距 15 磅 |
| 表格内容 | 跟随正文（默认宋体） | 跟随正文（宋体→Times New Roman） | 保留原字号 | 外边框 1.5 磅，自适应宽度 |

完整字段与可选项见 [`references/rules-schema.md`](references/rules-schema.md)，可直接编辑的示例见 [`references/default-rule.json`](references/default-rule.json)。

---

## 自定义排版规则

1. 复制 `references/default-rule.json` 为副本。
2. 按需修改（字体、字号、行距、页边距、跳过页数、表格边框宽度、是否外提标题等；字段含义见 `references/rules-schema.md`）。
3. 用 `--rules 副本.json` 传入。

规则 JSON 兼容插件导出的 `{ "Rules": [...], "ActiveRuleIndex": 0 }` 包装格式。

---

## 目录结构

```
设计报告排版skill/
├─ SKILL.md                 # 技能定义（Claude 读取，含调用与进度汇报流程）
├─ README.md                # 本文件
├─ CLAUDE.md                # 给 Claude Code 的仓库架构说明
├─ scripts/
│  ├─ FormatDocx.exe        # 排版命令行程序（已编译）
│  ├─ WordFormatter.dll     # 排版引擎（来自 WordTools 仓库）
│  ├─ Newtonsoft.Json.dll   # 依赖
│  └─ Microsoft.Office.Tools.Common.v4.0.Utilities.dll
├─ references/
│  ├─ rules-schema.md       # 规则字段完整说明
│  └─ default-rule.json     # 默认排版规则（水利工程报告）
├─ docs/
│  └─ 用户使用说明.html      # 面向终端用户的使用说明
└─ src/
   └─ FormatDocxHost/       # 命令行宿主源码（Program.cs + .csproj）
```

`tmp/` 为开发临时目录（测试文档、探查脚本、日志），不随技能分发。

---

## 开发与构建

命令行宿主源码在 `src/FormatDocxHost/`（.NET Framework 4.8）。**排版引擎本体不在本仓库**——`WordFormatter.dll` 来自同级仓库 `D:\Claude Code\WordTools\src\WordFormatter\`。修改宿主后重建：

```bash
# bash 下 /t /p 会被当路径转换，用双斜杠规避
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" \
  "src/FormatDocxHost/FormatDocxHost.csproj" //t:Rebuild //p:Configuration=Release
```

然后把 `src/FormatDocxHost/bin/Release/FormatDocx.exe` 复制到 `scripts/`。

> 若要改动**排版行为本身**（段落分类、标题识别、表格处理等），需在 `WordTools` 仓库修改 `WordFormatter` 源码、重建引擎，再把新的 `WordFormatter.dll` 复制到 `scripts/`。本仓库的 `Program.cs` 只负责启动 Word、加载规则、保存文件，以及表格字体「跟随正文」的解析。

---

## 注意事项

- 排版前确认目标文档**没有在任何 Word 窗口中打开**，否则会以只读打开并报错（退出码 3）。
- 网络下载/邮件附件的文档可能触发 Word「受保护的视图」，需先在 Word 中启用编辑。
- 大文档排版需几十秒到几分钟（COM 自动化逐段处理），属正常；超时设置不要低于 5 分钟。
- 排版会修改文档内容，重要文档请用 `--output` 另存或先备份。
- 表格标题外提仅处理**首行为合并单元格**的标题表；标题在非合并首行的表格不会被外提。

---

## 相关文档

- [SKILL.md](SKILL.md) — 技能定义与调用流程（含长文档进度汇报默认做法）
- [docs/用户使用说明.html](docs/用户使用说明.html) — 终端用户使用说明
- [references/rules-schema.md](references/rules-schema.md) — 排版规则字段说明
- [CLAUDE.md](CLAUDE.md) — 仓库架构与构建说明（面向 Claude Code）
