# 排版规则 JSON 字段说明（FormatRule）

规则文件为单个 FormatRule 对象（PascalCase 字段名）。完整示例见同目录 `default-rule.json`。

## 顶层结构

| 字段 | 类型 | 说明 |
|---|---|---|
| `Name` | string | 规则名称 |
| `Version` | string | 版本号（任意字符串） |
| `PageSetup` | object | 页面设置，见下 |
| `Styles` | object | 段落样式规则，键固定为 `heading1`~`heading4`、`body`、`figCaption`、`tblCaption` |
| `Table` | object | 表格规则，见下 |
| `Image` | object | 图片段落规则（字段同 StyleRule，仅对齐/行距生效） |
| `SkipPages` | int | 跳过前 N 页不排版（封面、扉页等），0 = 不跳过 |

## PageSetup

| 字段 | 类型 | 单位 | 说明 |
|---|---|---|---|
| `PaperSize` | string | — | 目前支持 `"A4"` |
| `MarginTop` / `MarginBottom` / `MarginLeft` / `MarginRight` | float | cm | 页边距 |

## Styles.* （StyleRule）

| 字段 | 类型 | 单位 | 说明 |
|---|---|---|---|
| `Label` | string | — | 显示名（仅备注用途） |
| `FontCN` | string | — | 中文字体，如 `"宋体"`、`"黑体"`、`"楷体"` |
| `FontEN` | string | — | 西文字体，如 `"Times New Roman"` |
| `FontSize` | float | pt | 字号（三号=16，四号=14，小四=12，五号=10.5） |
| `Bold` | bool | — | 加粗 |
| `Alignment` | string | — | `left` / `center` / `right` / `both`（两端对齐） |
| `SpaceBefore` / `SpaceAfter` | float | pt | 段前/段后间距 |
| `LineSpacing` | float | 见 LineRule | `LineRule=auto` 时为倍数（1.5 = 1.5 倍行距）；`exact`/`atLeast` 时为磅值 |
| `LineRule` | string | — | `auto`（多倍行距）/ `exact`（固定值）/ `atLeast`（最小值） |
| `FirstLineIndent` | float | pt | 首行缩进磅值（与 CharIndent 二选一） |
| `CharIndent` | int | 字符 | 首行缩进字符数（正文常用 2），优先于 FirstLineIndent |
| `OutlineLevel` | int | — | 大纲级别 1-9；10 = 正文文本。标题必须与其层级一致，否则导航窗格层级错乱 |
| `RemoveSpaces` | bool | — | 清理多余空格：删除中文之间的空格、段首段尾空格，保留英文单词间空格 |

## Table

| 字段 | 类型 | 说明 |
|---|---|---|
| `OuterBorderWidth` | float | 表格外边框宽度（pt），0 = 不设置 |
| `AutoFitWindow` | bool | 表格自适应窗口宽度 |
| `FontCN` / `FontEN` | string | 表格内中西文字体。**留空 = 跟随正文字体**（宿主层解析）：中文字体取正文 `body.FontCN`；西文字体在正文为宋体时取 `Times New Roman`，否则取 `body.FontEN`。显式指定（非空）优先。 |
| `FontSize` | float | 表格内字号（pt，五号=10.5），0 = 保留原字号 |
| `RemoveSpaces` | bool | 清理单元格内多余空格 |
| `ExtractTitles` | bool | 表格标题外移：把表格首行合并单元格中的标题提取到表外独立段落 |

## 标题识别规则（引擎内置，不可配置）

- `第X章/篇/部` → heading1；`第X节` → heading2
- `1`、`1.1`、`1.1.1`、`1.1.1.1` 数字编号 → 按深度 heading1~4（含 Word 自动编号）
- 已有 `标题 N` 样式或大纲级别的段落 → 按样式/级别识别
- 标题长度上限 60 字符
- `图...` / `表...` 开头且 ≤40 字符 → 图名/表名
- 表格内段落不参与段落格式化，仅受 Table 规则处理
