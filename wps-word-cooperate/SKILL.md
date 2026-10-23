---
name: wps-word-cooperate
description: 在 WPS Writer 里实时改 Word 文档——打开用户正开着的那份 .docx，以修订模式逐段改写（改动是可见的增删标记，可逐条接受/拒绝）、加批注、把页面渲染成图核对版式，所有改动当场出现在 WPS 窗口里。当用户要求：在 WPS 里改这个文档、在 WPS 里做修订、实时改 Word、一边看着一边改、在文档里加批注、核对排版/版式，或提到 wps word、live edit word 时使用。不适用：纯离线批量处理（改副本、不要求用户在场，用 docx-paragraph-polish 等离线技能）、用户明确不要动原文件（只出旁挂报告）、Excel/PPT（本 skill 只做 Word）。
---

# wps-word-cooperate：在 WPS Writer 里实时改 Word

**能力边界（一句话）**：只做 **Word/.docx**、只走 **WPS Writer**（ProgID `KWPS.Application`，本机没装微软 Office 也照跑）；所有文字改动都是**修订（track changes）**——用户能逐条接受或拒绝，不存在"悄悄改掉原文"这条路。

> 下文 `<skill>` = 本 skill 的安装目录（如 `C:\Users\Administrator\.dsh\skills\wps-word-cooperate`）。命令一律按**完整可跑**形式给出：
> `powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\wps.ps1 <命令> [参数]`
> （装了 PowerShell 7 的机器可以把 `powershell` 换成 `pwsh`，参数完全一样。**本机没有 `pwsh`，照上面写就行**。）
> 输出只有一段 JSON：`{"ok":true,"result":{...}}` 或 `{"ok":false,"error":{"code":"...","message":"..."}}`。
> 环境：Windows + WPS；首次调用会用系统自带 `csc.exe` 编译 helper（约 1 秒，之后按源码哈希缓存，不再编译）。

## 铁律（五条，一条都不能省）

1. **先 `status`，再动手，并说出在哪儿动手。** 每次开工第一条命令是 `status`，把"几个实例在跑 / 开着哪些文档 / 哪份是 active / 有没有未保存修改"念给用户，并明确报出**即将修改的那份文档的完整路径**。不碰用户没指定的文件。
2. **`readOnly` 就是不许改。** `open`/`read` 返回 `readOnly: true`（文件被别处占用或有只读属性）时，`revise`/`comment` 会直接报 `READONLY`。**不要**用另存、改文件属性、杀占用进程等方式绕过去——先问用户。
3. **不替用户关文档。** `close` 只关本次明确要关的那一份；用户开着的其它文档、用户自己的 WPS 窗口一律不动。关之前先 `status` 确认这份是本次打开的。
4. **结束必须 `close`，并汇报 `wps.exe` 数量。** 收尾命令 `close -Doc <路径>` 会返回 `processes`（当前 wps.exe 计数）、`documentsLeft`。把这两个数报给用户；只清理自己起的无窗口 `/Automation` 实例，**绝不**按映像名批量杀 `wps.exe`。（注：WPS 退出后会自己预拉起 1–2 个隐藏进程，约 1 分钟自灭，别把它当成泄漏。）
5. **交付前提醒"本机可能落盘即密文"。** 装了单位透明加密的机器上，WPS `save` 之后文件头变成 `17DA5FA01633CD9A`，直接发给同事打不开。保存后必须提这一句，并按 [reference/encryption.md](reference/encryption.md) 处理。

## 命令速查

| 命令 | 参数 | 作用 |
|---|---|---|
| `doctor` | — | 自检：哪个套件可用/版本/是否在跑/开着的文档/渲染是否可用/各类进程数 |
| `status` | — | 在跑的实例与文档清单（name / path / saved / readOnly / active / paragraphs） |
| `open` | `-Doc <路径> [-Show] [-ReadOnly]` | 已开则复用，未开则打开；返回 readOnly、段落数、页数 |
| `read` | `-Doc <路径> [-From 1] [-Count 60] [-Tables]` | 段落（编号/样式/自动编号/文本/inTable）+ 表格清单 + 页数/修订数/批注数；表格清单默认就返回 |
| `revise` | `-Doc <路径> -Items <json文件>` | **修订模式**批量改写（见下），返回改动与修订计数 |
| `comment` | `-Doc <路径> -Para N [-Anchor "文字"] (-TextFile <txt文件> \| -Text "短文本")`；删除：`-Delete -Index N \| -All` | 加批注（中文正文走文件，避开引号转义） |
| `render` | `-Doc <路径> -Page 1 [-To 4] -Out <目录> [-Width 1100]` | 页面 EMF → PNG（宽 320–2400），返回每页路径与尺寸 |
| `font` | `-Doc <路径> -Name <字体名> [-Track]` | 把**正文、页眉、页脚、文本框**里的字体统一改成指定字体（中英文一起改，`-Name 仿宋`）；返回每个区段的结果与回读值。默认**不产生修订**，整批算一个撤销单元 |
| `format` | `-Doc <路径> [-Para N \| -From N -To M \| -Style 样式名]` + 任意格式项 | 对**整篇/一段/一段区间/某样式的所有段**套字符与段落格式（见下） |
| `style` | `-Doc <路径> -Style <样式名> [-NoNumbering]` + 任意格式项 | 改**样式本身**（正文、标题 1…），所有用它的段落一起变 |
| `page` | `-Doc <路径> [-Paper A4] [-Orientation portrait\|landscape] [-Top/-Bottom/-Left/-Right 厘米] [-HeaderText] [-FooterText] [-PageNumber center\|right\|left\|none] [-PageNumberStart N] [-All]` | 纸张/方向/页边距/页眉页脚/页码（逐节应用） |
| `table` | 插入：`-Para N -Data <json> [-Borders grid\|three-line\|none] [-Widths "2,3,2"] [-HeaderRow] [-Caption]`；改表：`-Table n [-Borders ...] [-Autofit] [-RowHeight] [-KeepTogether] [-Font/-Size/-Align]` | 建表 / 改表（框线、三线表、表头跨页重复、行不跨页、列宽、行高） |
| `caption` | `-Para N -Label 表\|图 -Title "..." [-Above] [-Align center]` | 生成**域**式题注「表 1 xxx」（自动编号，增删后重排） |
| `picture` | 插图：`-Path <图片> -Para N [-Width "12cm"] [-Caption] [-Align]`；调整：`-Image n [-Width/-Height] [-CropLeft/-CropRight/-CropTop/-CropBottom 百分比]` | 插入图片（可带图题）/ 调尺寸与裁剪 |
| `toc` | `-Doc <路径> [-Para N] [-Title 目录] [-Levels 3]` | 插入并刷新目录（域），可从文首或指定段之前 |
| `reference` | 列表：`-Action list -Items <json 数组> [-AtPara N] [-Heading 参考文献]`；引用：`-Action cite -Para N -Numbers "1,2,5"\|"3-6"` | 生成参考文献列表 / 在段末插入上标引用 |
| `view` | `-Doc <路径> -Para N` 或 `-Page N` | 把 WPS 窗口滚到那一段/那一页（唯一能做的"跟随"，不动文档内容） |
| `fields` | `-Doc <路径>` | **刷新全部域**：页码、目录、题注编号、引用（改完正文、交付前跑一次） |
| `cell` | `-Doc <路径> (-Table n -Row r -Col c \| -Para N) -Text "..."` + 任意格式项 | 改一个表格单元格的文字（可带字体字号对齐） |
| `insert` | `-Doc <路径> [-Para N] [-Where after\|before\|end] -Items <json> [-Track]` | **批量插入段落**（默认干净写入、不进修订；`-Track` 才走修订），可带样式与格式 |
| `revisions` | `-Doc <路径> [-Action list\|accept\|reject] [-Index N \| -All]` | 列出**未处理**的修订（类型/作者/所在段/文本），或按人要求接受/拒绝 |
| `snapshot` | `-Doc <路径> [-Tag 名字]` | 存一份段落文本快照（给 `diff` 用，存在 %LOCALAPPDATA%，不污染文档目录） |
| `diff` | `-Doc <路径> [-Since 快照路径\|last]` | 和快照比：哪些段被改了（before→after）、哪些增删了——**人改的和 Agent 改的都算** |
| `handoff` | `-Doc <路径> [-Note "本轮做了什么"] [-As 输出路径]` | 在文档旁生成 `.handoff.md` 交接单：状态/待办/操作记录 |
| `save` | `-Doc <路径> [-As <另存路径>]` | 保存 / 另存 |
| `close` | `-Doc <路径> [-Save]` | 关闭；无其它文档时退出实例，并回报进程数 |
| `quit` | — | 关掉该实例的**全部**文档并退出（不保存）。**慎用**：会波及用户在同一实例里开的文档，只在自己的独占实例上用 |

所有命令都可用 `-TimeoutSec <秒>`（默认 180）控制等待；超时返回 `TIMEOUT`，说明 WPS 大概率弹了对话框。

### `revise` 的 items 格式

```json
{"items":[
  {"para":12,"expect":"3.2 工程地质","op":"replace","find":"孔口高程","replace":"孔口地面高程"},
  {"para":15,"expect":"根据地质勘察","op":"set_text","text":"根据地质勘察成果，坝址区基岩为……"},
  {"para":20,"expect":"4.1 施工组织","op":"insert_after","text":"4.1.1 施工条件"},
  {"para":31,"expect":"本阶段不再","op":"delete"}
]}
```

- `op` 四种：`replace`（字符级，**优先用**）、`set_text`（整段删+插）、`insert_after`、`delete`。
- `expect` 必填，校验该段文本是否**以它开头**，不符即 `ANCHOR_MISSING`（防改错位置）；`replace` 的 `find` 找不到也报同一码。
- `replace` 默认只替换该段首个命中；同一段内要多处替换，加 `"all":true`。
- 批内自动按 `para` **从大到小**执行，编号不会互相串；要插多段时按原始编号一次写全。
- **整批是原子的**：动手前先把每一条都核一遍（段落存在、`expect` 对得上、`find` 找得到），有一条不过就**一条都不改**并报 `ANCHOR_MISSING`；万一执行中途出错，已产生的修订会被逐条 `Reject` 回滚，错误信息里会写"rolled back N revision(s)"。
- 每条结果带 `textAfter`（该段改完后的开头 90 字）——**立刻比对它**，能第一时间发现"替换后变成「进行了了」"这类找错位置的问题，不必再 `read`。
- 返回 `applied`、`revisionsBefore`、`revisionsAfter`、`revisionsAdded`。
- **段落编号的坑**：表格里每个单元格算一段，且每行末尾**再多算一段**（3×4 表占 15 个段落号），所以表格附近的段落号不能按"行×列"推算；`read` 会给表内段落打 `inTable: true`。

## 典型工作流

### a) 读一遍并核对版式

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\wps.ps1 status
powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\wps.ps1 open -Doc "D:\报告\某报告.docx" -Show
powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\wps.ps1 read -Doc "D:\报告\某报告.docx" -From 1 -Count 120
powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\wps.ps1 render -Doc "D:\报告\某报告.docx" -Page 1 -To 3 -Out "D:\报告\_pages"
```

看图时按这几项过：表格是否出框/掉到下一页、图题表题是否贴着自己的图表、有没有整段空行或半空页、修订红字与批注气泡是否在预期段落。（`render` 用的是打印视图的页面图元，**不是**截图，窗口在后台也出图。）

### b) 逐段润色（`replace` 优先）

1. `read -From N -Count 40` 分批读，拿到段落编号与原文——编号是唯一锚点。
2. 造 items：能 `replace` 就 `replace`（只产生字符级删/插，最贴近润色）；**只有整段重构才用 `set_text`**（会产生整段删+插，用户不好逐字看）。
3. items 写成 UTF-8 的 JSON 文件，`revise` 一次提交。
4. `render` 抽 1–2 页看修订落点（改动后 WPS 视图会跳回光标，别指望"跟着滚"）。
5. 用户看过再 `save`；保存后按铁律 5 提醒密文。

### c) 校审批注（`comment` + `render` 复核）

1. `read` 定位问题段，记下编号与该段开头文字。
2. 每条批注写一个 UTF-8 的 txt（`-TextFile`，避免命令行引号转义把中文弄坏），`-Anchor` 指到具体词句可让批注贴在那几个字上。
3. 批注正文写"问题 + 依据 + 建议"，别写"见上"。作者固定为当前登录用户，要区分来源就在正文里署名（如"（AI 校对）"）。
4. `render` 复核气泡是否落在正确段落（页面边上，跨页易错位），再 `save`。

### d) 统一字体（`font`）

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File <skill>\scripts\wps.ps1 font -Doc "D:\报告\某报告.docx" -Name 仿宋
```

- 覆盖**正文（含表格）、各节页眉页脚、浮动文本框**；中英文与符号一起改（`NameFarEast`/`Name`/`NameAscii`/`NameOther` 都设）。
- 只改**现有文字**：以后新打的字仍随段落样式的字体。要让后续输入也变，得改样式本身（本命令暂不做，用 WPS 的"修改样式"更稳妥）。
- 默认**不进修订**——整篇字体改动若逐字留痕会有成千上万条格式修订，把文字修订淹没；整批包在一个撤销记录里，**按一次 Ctrl+Z 全部撤回**。
- 返回 `ranges`（每个区段的区段名、字符数、成败）与 `readBack`（回读正文的东亚字体名，应等于你给的字体名）。回读为空/不符说明没落上，先别保存。
- 字体名用 Windows 显示名（`仿宋`、`黑体`、`宋体`、`楷体`）；英文名（`FangSong`）在部分环境不认。

## 格式项（`format` / `style` / `table` / `caption` 通用）

| 参数 | 取值 | 说明 |
|---|---|---|
| `-Font` | `仿宋` | 中英文一起换（`NameFarEast`/`Name`/`NameAscii`/`NameOther`） |
| `-LatinFont` | `"Times New Roman"` | 只换西文，不动中文字体 |
| `-Size` | `12` / `"10.5pt"` | 字号（磅） |
| `-Bold` `-Italic` `-Underline` | 开关 | 关掉用 `-Bold:$false` |
| `-Super` `-Sub` | 开关 | 上标 / 下标 |
| `-Color` | `"#C00000"` | 字体颜色 |
| `-Align` | `left\|center\|right\|justify` | 对齐（中文也认 `居左/居中/居右/两端对齐`） |
| `-IndentChars` | `2` | **左缩进 2 字符**（中文单位，比磅值可靠） |
| `-LeftIndent` | `"0.74cm"` / `"21pt"` | 左缩进（长度） |
| `-FirstLineIndent` | `2` / `"0.74cm"` / `0` | **首行缩进**：纯数字=字符数（2=两字），带单位=长度，0=取消 |
| `-LineSpacing` | `1.5` / `"20pt"` / `"at least 20pt"` / `single` / `double` | 数字=倍数，`pt`=固定值，`at least`=最小值 |
| `-SpaceBefore` `-SpaceAfter` | `"6pt"` / `0` | 段前 / 段后间距 |
| `-PageBreakBefore` | 开关 | 该段另起一页 |
| `-Track` | 开关 | 让**格式改动也进修订**（默认不进；整篇排版不要开） |

### e) 一键排版（`format` + `style` + `page`）

```powershell
<skill>\scripts\wps.ps1 status
# 全文正文：仿宋小四、首行缩进 2 字符、1.5 倍行距、两端对齐
<skill>\scripts\wps.ps1 format -Doc "D:\报告\某报告.docx" -Font 仿宋 -Size 12 -FirstLineIndent 2 -LineSpacing 1.5 -Align justify
# 一级标题：黑体三号、加粗、段前后距、左对齐、不分首行缩进、段前分页
<skill>\scripts\wps.ps1 style -Doc "D:\报告\某报告.docx" -Style "标题 1" -Font 黑体 -Size 16 -Bold -Align left -FirstLineIndent 0 -SpaceBefore "12pt" -SpaceAfter "6pt"
# 页面：A4、页边距、页脚居中页码
<skill>\scripts\wps.ps1 page  -Doc "D:\报告\某报告.docx" -Paper A4 -Top 2.54 -Bottom 2.54 -Left 3.17 -Right 2.54 -PageNumber center
# 复核
<skill>\scripts\wps.ps1 render -Doc "D:\报告\某报告.docx" -Page 1 -To 3 -Out "D:\报告\_pages"
```

要点：
- **先 `style` 后 `format`**：样式管"以后新写的字"，`format` 管"现在已经写好的字"，两者一起上才不返工。
- `format` 的三种范围：`-Para N`（一段）、`-From N -To M`（区间）、`-Style "标题 2"`（该样式全部段落）；**不给范围 = 全文**。
- 整篇排版不要开 `-Track`：会生成海量格式修订。排版类改动默认不进修订，一次 Ctrl+Z 全撤。
- 排版前后各 `render` 一次，重点看：标题是否掉行、首行缩进是否一致、表格是否被挤到下一页、页码是否只有一节有。

### f) 表格、题注、图片（`table` / `caption` / `picture`）

```powershell
# 建表：数据走 JSON 文件（UTF-8），三线表 + 表头跨页重复 + 表题
<skill>\scripts\wps.ps1 table -Doc "D:\报告\某报告.docx" -Para 120 -Data "D:\报告\表1.json" -Borders three-line -HeaderRow -Caption "主要水文特征值"
# 改已有表：全框线、按窗口自适应、行不跨页
<skill>\scripts\wps.ps1 table -Doc "D:\报告\某报告.docx" -Table 1 -Borders grid -Autofit window -KeepTogether
# 给已有表补题注（域，自动编号）
<skill>\scripts\wps.ps1 caption -Doc "D:\报告\某报告.docx" -Para 120 -Label 表 -Title "主要水文特征值" -Above
# 插图（带图题，图题在下）
<skill>\scripts\wps.ps1 picture -Doc "D:\报告\某报告.docx" -Path "D:\报告\图1.png" -Para 90 -Width "12cm" -Caption "工程位置示意"
```

- `表1.json` 形如 `{"data":[["断面名称","集水面积(km²)"],["白沙","1560"]]}`；也接受不带外层对象的裸数组。
- `-Borders three-line` = 论文三线表（上下粗、表头下细、其余无线）；`grid` 全框线；`none` 无线。
- 题注是 **SEQ 域**（`表 1`、`图 1` 自动编号，增删后重排），不要手打编号——手打的不会重排。
- 表题在**表上方**、图题在**图下方**，这是规范；`caption` 有 `-Above` 控制，`table -Caption` / `picture -Caption` 已按此默认。
- 图片宽度用长度（`"12cm"` / `"340pt"`）；插图后**必须 `render` 看**是否出框、是否被压扁。

### g) 目录与引用（`toc` / `reference` / `view`）

```powershell
# 文首插目录（一级到三级），自动刷新
<skill>\scripts\wps.ps1 toc -Doc "D:\报告\某报告.docx" -Title 目录 -Levels 3
# 参考文献列表（JSON 字符串数组）
<skill>\scripts\wps.ps1 reference -Doc "D:\报告\某报告.docx" -Action list -Items "D:\报告\refs.json" -Heading 参考文献
# 在正文某段末加引用上标 [1,2,5]
<skill>\scripts\wps.ps1 reference -Doc "D:\报告\某报告.docx" -Action cite -Para 208 -Numbers "1,2,5"
# 让 WPS 窗口滚到第 208 段
<skill>\scripts\wps.ps1 view -Doc "D:\报告\某报告.docx" -Para 208
```

- 目录**依赖标题样式**：文档里没有「标题 1/2/3」的段落，目录就是空的——先 `style` 或在 WPS 里给标题套样式。
- 目录是域，插完要刷新；`toc` 已内置 `Update()`。改完正文后再跑一次 `toc` 即可更新页码。
- `view` 只滚窗口、不改文档，是"跟随修改位置"能做到的最大程度（不是插件那种跟随滚动）。
- 拉到最后要提醒：**目录、页码、域在 WPS 里刷新过再交给用户**——用 `fields` 一条命令刷新全部域，比逐项找更省事。

### j) 和人接力（`snapshot` / `diff` / `revisions` / `handoff`）

**这是本 skill 的核心用法**：Agent 接管一段 → 人裁决/手改 → Agent 读人的结果继续。四个动作：

```powershell
# ① 开工前先拍快照（记下"我接手时文档长什么样"）
<skill>\scripts\wps.ps1 snapshot -Doc "D:\报告\某报告.docx" -Tag round1
# ② 干活……（revise / format / table …）
# ③ 人介入之后：看他改了什么、还剩多少修订没裁决
<skill>\scripts\wps.ps1 diff      -Doc "D:\报告\某报告.docx" -Since round1
<skill>\scripts\wps.ps1 revisions -Doc "D:\报告\某报告.docx"
# ④ 收尾：写一份人看得懂的交接单在文档旁边
<skill>\scripts\wps.ps1 handoff -Doc "D:\报告\某报告.docx" -Note "本轮润色 12 段，其中 3 段改了数据表述，待你确认"
```

- **`snapshot`**：存每段文本到 `%LOCALAPPDATA%\wps-word-cooperate\snapshots\`（**不写进用户文档目录**）；`-Tag` 起名，`last` 指最新一份。
- **`diff`**：和快照比，给出 `changed`（某段 before→after）/ `added` / `removed` / `same`。**人改的和 Agent 改的都在里面**。
- **`diff` 比的是"接受视图"**（WPS 的 `Range.Text` 就是接受后的样子：插入计入、删除不计）。所以：人**接受**了 Agent 的修订 → 文本本来如此，`diff` 看不到变化（剩余量看 `revisions`）；人**拒绝**或手改 → `diff` 报成 `changed`。判断"还有多少待裁决"永远看 `revisions`，判断"文本变成什么样"看 `diff`。
- **`revisions`**：列出**未处理**的修订（类型/作者/时间/所在段/文本）。已被人接受或拒绝的不在列表里，所以"列表"天然就是"待办"。
  - `-Action accept -Index N` / `-Action reject -Index N` 处理单条；`-All` 处理全部。
  - **纪律**：`-All` 只在用户明确说"都接受/都拒绝"时才用；否则一律让人自己在 WPS 里裁决。
- **`handoff`**：在文档旁生成 `<文档名>.handoff.md`——文档状态、待处理修订/批注数、这一轮做了什么（你传的 `-Note`）、最近 30 条操作记录、"接手时建议先做"。下一轮 Agent 或同事直接读这份就能接上。
- **`changedOutside`**：任何写命令的结果里若出现这个字段（还有 `changedOutsideHint`），说明**在你上次写入之后，文档被人动过**。此时不要闷头继续：先 `diff` 看变了什么、`revisions` 看还剩什么，再决定。

**两套接力剧本**

*批次式（A）*：`snapshot` → 批量改（修订）→ 请人处理 → `diff` + `revisions` 读结果 → 第二轮 → `save` → `handoff`。
原则：**一轮只做一类事**（润色 / 排版 / 校审分开），人处理完再开下一轮；每轮开头都 `snapshot`。

*并发式（B）*：人和你同时在看同一份文档。
原则：① 写之前先看 `changedOutside`；② `revise` 的 `expect` 报 `ANCHOR_MISSING` 时，先 `read` 那一段——**很可能是人改了它**，不要当成编号漂移盲目全量重读；③ 不替人裁决修订；④ 人的光标位置别去动（不要 `Select`，`view` 默认也不动光标）。

### i) 刷新域 / 改单元格 / 干净插入（`fields` / `cell` / `insert`）

```powershell
# 改完正文、存盘之前：刷新页码、目录、题注、引用
<skill>\scripts\wps.ps1 fields -Doc "D:\报告\某报告.docx"
# 改一个表格单元格（带格式）
<skill>\scripts\wps.ps1 cell -Doc "D:\报告\某报告.docx" -Table 1 -Row 1 -Col 2 -Text "集水面积(km²)" -Font 宋体 -Size 10.5 -Align center
# 一次插入多段（干净写入，不是修订），带样式与格式
<skill>\scripts\wps.ps1 insert -Doc "D:\报告\某报告.docx" -Para 40 -Where after -Items "D:\报告\新增.json"
```

- `insert` 的 items：`[{"text":"第一段","style":"正文","font":"仿宋","size":12,"firstLineIndent":2}, ...]`；`-Text` 可直接插一段。
- **插入位置会受表格影响**：锚点段落紧挨表格（或本身就是单元格）时，新段落会落进表格单元格里——返回里 `inTable: true` 就是提示；要插在表格前面，改用 `-Where before` 指向表格之后的那一段，或改完再 `read` 确认。
- `insert` 默认**不进修订**（写新章节、填模板单元格用），要留痕加 `-Track`。
- `cell` 的 `-Text ""` 清空单元格；合并单元格会报 `BAD_ARGS`（WPS 的 `Cell(r,c)` 取不到）。
- 这三条都不保存；`fields` 只刷新显示，不改文字。

### h) 排版实测坑（照做少返工）

1. **`table -Para` 必须指向正文段**，不能指到表格内部：`read` 里带 `inTable: true` 的段落会报 `BAD_ARGS`（要往单元格里塞表就加 `nested: true`，且本 WPS 的嵌套表时灵时不灵，报错时改用"把表放在表格之后"）。
2. **题注是段落在前、表格在后**：`table -Caption` 会自动把「表 N 标题」放在表格上方一段；`picture -Caption` 放在图片下方。别再手打「表 1」。
3. **图片裁剪参数是百分比 0–100**（`-CropLeft 25` = 裁掉左边 25%）；WPS 内部的 `Crop*` 单位是"原图 100% 尺寸的磅"，换算已在命令里处理，别自己再换一次。
4. **`toc` 默认插在真正的文首**（封面之前）。要放在封面/扉页之后，给它 `-Para <正文第一段的编号>`；插之前先确认文档里有「标题 1/2/3」样式，否则目录会是空的（返回 `entries:0` + 提示）。
5. **加粗在渲染图里可能看不出来**：WPS 的 EMF 页面图对 `Bold` 的呈现不可靠（属性、docx 里的 `w:b` 都是对的）。别因为截图里"看着没粗"就反复重设；要看加粗去 WPS 窗口里看，或检查 `read` 的样式。
6. **表格跨页**：`-HeaderRow`（表头重复）+ `-KeepTogether`（行不跨页）一起用；长表就是会跨页，别硬压。
7. **先看文档原来怎么编页码**：中文报告常在**页眉**放页码（奇偶页还可能不同）。`page -PageNumber center` 是往**页脚**插 PAGE 域，文档本来就有页眉页码时会变成两套。动手前先 `status`/`render` 看一眼，或直接不传 `-PageNumber`。
7.5. **紧挨批注的文字，位置会差 1 个字符**：WPS 里批注锚点占一个位置、却不出现在 `Range.Text` 里，所以对着批注旁边的字做 `replace` 会错位（实测把"对无实测"改成了"对无实测测"）。现在 `replace` 会**先读回目标片段校验**，位置不对时在 ±3 字符内自动纠正，纠正不了就报 `ANCHOR_MISSING` 而**不会写坏**。确实要改那段：`comment -Delete -Index n` 摘掉批注 → 改 → 再 `comment` 按原词补回。
8. **排版类命令都不进修订**（`format`/`style`/`page`/`table`/`caption`/`picture`），整批一次 Ctrl+Z 可撤；`revise`/`comment` 才是走修订的那两个。

## 与其它 skill 的关系

- **`office_*` 工具（dsh-office 插件）**：装了那个插件时，`office_open/office_read/office_edit/office_render` 与本 skill 能力重叠、互不冲突——插件适合"顺手改一份"，本 skill 适合"要分发给同事、不想让每人装插件"的场合。两者都可用时**选一个用到底**，别交替（会各开实例）。
- **`docx-paragraph-polish`**（离线润色，纯 OpenXML 写 `w:ins`/`w:del`，产出副本）：**本机有 WPS 且用户要在场逐条过**时，改用本 skill——修订直接出现在 WPS 窗口里、可当场接受/拒绝、顺手能核对版式，不必先生成副本再搬家。反之，无人值守、上千段一次跑完的批量润色仍用 `docx-paragraph-polish`（它不依赖 WPS，也不受"落盘即密文"影响）。
- **`design-report-review`**（校审）：先跑校审拿到问题清单，再用本 skill 的 `comment` 把结论逐条落进原文对应段落，最后 `render` 复核——用户拿到的不是旁挂的 HTML，而是**带批注的原稿**。

## 故障排查

| 错误码 / 现象 | 怎么办 |
|---|---|
| `NO_APP` | 没装 WPS 或 COM 没注册。先手动打开一次 WPS Writer 再重试；`doctor` 里 `apps.wps.installed=false` 说明这台机器确实没有 WPS。 |
| `NOT_OPEN` | `-Doc` 与实际打开的文档不是同一路径（另存过/短路径/大小写/相对路径）。先 `status`，把 `path` 原样复制过来。 |
| `READONLY` | 文件被别处占用或是只读属性。关掉占用它的程序，或请用户另存一份再改；不要强行覆盖保存。 |
| `ANCHOR_MISSING` | 段落编号漂移（表格也占段落号，且每行多一段）或 `expect`/`find` 与原文不符。整批**一条都没改**，重新 `read` 该段核对编号与原文再来。 |
| `BUSY` / `TIMEOUT` | 用户正在打字、或有对话框/弹窗挡着（等满 `-TimeoutSec` 秒即超时）。等几秒或请用户关掉对话框再重试，别连发。 |
| 打开即 `COM_ERROR` / 文件读不出来 | 多半是密文且不归 WPS 管（例如离线脚本写出的）。让用户在 WPS 里打开一次（驱动会透明解密），或按 `decrypt-office-pdf` 先解密。注意：helper **不产出** `ENCRYPTED` 码。 |
| 中文乱码 | 控制台代码页不是 UTF-8。先 `[Console]::OutputEncoding=[Text.Encoding]::UTF8`，或把 JSON 落成文件按 UTF-8 读。 |
| `wps.exe` 变多 | 中途异常没 `close`，或别处 `New-Object` 开了新实例。先 `status` 看实例数，只结束无窗口的 `/Automation` 实例，用户窗口留着，然后汇报数字。 |

## 参考

- [reference/wps-quirks.md](reference/wps-quirks.md) —— WPS 与 Word 的实测差异表（冒充版本号、样式中文名、`Range` 是方法、关最后一个文档静默杀实例……）。
- [reference/encryption.md](reference/encryption.md) —— 落盘即密文的现象、为什么 COM 路线不受影响、交付与解密链路。
