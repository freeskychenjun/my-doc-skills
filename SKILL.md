---
name: design-report-format
description: 一键排版 Word 设计报告（.doc/.docx）——页面设置、一至四级标题、正文、图名/表名、表格边框与自适应宽度、多余空格清理。当用户要求对 Word 报告/文档进行排版、格式化、统一字体字号行距，或处理水利工程报告等工程技术文档排版时使用。
---

# 设计报告一键排版

通过 `scripts/FormatDocx.exe` 调用 Word COM 自动化 + WordFormatter 排版引擎，对 Word 文档执行完整排版流水线。

## 运行环境

- Windows + 桌面版 Microsoft Word 2010 及以上（不支持 WPS、Word Online）
- 无需安装 VSTO 插件，脚本自带全部依赖（scripts/ 下的 DLL）

## 用法

```bash
scripts/FormatDocx.exe <输入.docx> [--rules 规则.json] [--output 输出.docx]
```

- 不带 `--output`：原地保存（覆盖原文件）
- 带 `--output`：另存为新文件（推荐，保留原稿）
- 不带 `--rules`：依次尝试 `%AppData%\CJTools\rules.json`（WordTools 插件的规则）、内置默认规则（水利工程报告）

成功时 stdout 输出一行 JSON 统计：

```json
{"ok":true,"file":"...","rule":"水利工程报告","paragraphs":27,"headings":7,"body":6,"images":0,"figCaptions":1,"tblCaptions":1,"extractedTableTitles":0,"skipped":12,"elapsedMs":602}
```

退出码：0 成功；1 排版出错；2 参数/规则文件错误；3 文档被占用、只读或受保护视图。

## 长文档进度汇报（默认做法）

`FormatDocx.exe` 全程往 **stderr** 输出进度（每处理 50 段一行，含阶段、当前/总数、已用时），**stdout 只在结束时输出一行 JSON 统计**。排版真实报告（通常较长）时，**默认按下法操作，让用户看到进度**：

1. **后台启动**，stderr 重定向到日志文件：
   ```bash
   scripts/FormatDocx.exe <输入.docx> [--rules 规则.json] [--output 输出.docx] > 排版.log 2>&1
   ```
2. **每隔约 20–30 秒** `tail` 一次日志，把最新进度行转述给用户，例如：
   - `scan 1500/2899 已用时 12s`（扫描分类中）
   - `format 250/357 标题 24 正文 198 图片 12 图标 11 表标 5 | 已用时 19s`（逐段排版中，带各类实时计数）
3. **反复轮询**，直到日志末尾出现 `{"ok":true,...}` 那行 JSON 即完成，再把统计（标题/正文/图名/表名/外提标题数/耗时）汇总给用户。
4. 若退出码非 0（尤其 **3 = 文档被占用/只读/受保护视图**），按退出码含义提示用户处理（关闭 Word 窗口、启用编辑等）。

阶段含义：`scan` 扫描分类（长文档主要耗时在此）、`format` 逐段排版、`table` 表格排版（一次性，无分项）、`page`/`extractTitles`/`listFont` 较快。

> 很小的测试文档（几十段、几秒完成）可直接前台跑，无需分段汇报。

## 排版内容

1. 页面设置（页边距、纸张方向）
2. 表格标题外移（默认开启，`Table.ExtractTitles=true`；仅处理首行为合并单元格的标题表）
3. 段落分类与格式化：一至四级标题（"第X章"/"1.1"/"1.1.1.1" 等自动识别）、正文、图名、表名、图片
4. 列表编号字体
5. 表格格式化：外边框、自适应窗口宽度、字体、多余空格清理

## 定制排版规则

1. 读 `references/rules-schema.md` 了解全部字段含义
2. 复制 `references/default-rule.json` 为副本，按需修改（字体、字号、行距、页边距、跳过页数等）
3. 用 `--rules 副本.json` 传入

规则 JSON 也兼容 WordTools 插件导出的 `{ "Rules": [...], "ActiveRuleIndex": 0 }` 包装格式。

## 注意事项

- **排版前先确认目标文档没有在任何 Word 窗口中打开**，否则会以只读打开并报错（退出码 3）
- 网络下载/邮件附件的文档可能触发 Word「受保护的视图」，需用户先在 Word 中启用编辑
- 大文档排版需几十秒到几分钟，属正常（COM 自动化），超时设置不要低于 5 分钟
- 排版会修改文档内容，重要文档务必用 `--output` 另存，或提醒用户先备份
- 图名/表名识别有 40 字符上限，超过的段落按正文处理（引擎内置行为）

## 重建宿主程序（一般不需要）

宿主源码在 `src/FormatDocxHost/`，排版引擎复用 WordTools 项目的 `WordFormatter.dll`。修改后用 MSBuild Release 重建并把产物复制到 `scripts/`：

```bash
MSBuild.exe src/FormatDocxHost/FormatDocxHost.csproj -t:Rebuild -p:Configuration=Release
```
