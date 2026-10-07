# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Claude Code **skill** (`design-report-format`, defined in `SKILL.md`) that one-click typesets Word design/engineering reports — page setup, 1–4 level headings, body text, figure/table captions, and table borders — via Word COM automation. It is the packaging + CLI host around a formatting engine that lives in a *separate* sibling repo.

This repo owns only:
- the **skill definition** (`SKILL.md`, `references/`, `docs/`),
- the **CLI host** source (`src/FormatDocxHost/`) and the compiled artifact shipped at `scripts/FormatDocx.exe`,
- the **bundled dependencies** in `scripts/`.

It does **not** own the formatting engine. That is `WordFormatter.dll` (namespace `WordFormatter.Engine` / `WordFormatter.Models`), built from `D:\Claude Code\WordTools\src\WordFormatter\`. The host's `.csproj` references it via `HintPath ..\..\..\WordTools\...`. See "Changing formatting behavior" below.

## Common commands

### Run the formatter
```bash
scripts/FormatDocx.exe <input.docx> [--rules rules.json] [--output out.docx]
```
- No `--output` → saves **in place** (overwrites). The SKILL runner should normally pass `--output` (e.g. a `-排版`-suffixed copy) to preserve the original.
- No `--rules` → rule load precedence: `--rules` file → `%AppData%\CJTools\rules.json` (the WordTools plugin's rules) → `FormatRule.CreateDefault()` built-in water-engineering default.
- Needs a timeout **≥ 5 minutes**; large reports run for minutes.

### Rebuild the host exe (rarely needed)
The .csproj hard-depends on the sibling WordTools Release build, so build WordTools first (or ensure its `bin\Release\WordFormatter.dll` + `Newtonsoft.Json.dll` exist), then:
```bash
MSBuild.exe src/FormatDocxHost/FormatDocxHost.csproj -t:Rebuild -p:Configuration=Release
```
Then copy the new `FormatDocx.exe` (and any updated DLLs) into `scripts/`. Target is .NET Framework 4.8.

### Inspect a formatted doc (dev, not shipped)
`tmp/` is scratch for throwaway test docs, logs, and PowerShell probes. To verify formatting, write a one-off `tmp/*.ps1` that opens the doc via Word COM (`New-Object -ComObject Word.Application`, read-only) and dumps `$p.Style.NameLocal`, `$p.Format.OutlineLevel`, `$p.Range.Font` (`.NameFarEast`=CN font, `.NameAscii`=Western, `.Size`), and `$p.Range.Information(12)` (wdWithInTable). Write results via `[IO.File]::WriteAllText(..., UTF8)` and Read them back — don't pipe Chinese through the console.
**PowerShell gotcha:** a `.ps1` saved UTF-8 *without BOM* is read as the system ANSI codepage (GBK), so Chinese **literals inside the script** arrive mojibake'd. Pass Chinese as CLI `-param` values (those arrive UTF-8) or `Get-Content -Encoding UTF8` from a file — never inline in a no-BOM `.ps1`.

## Architecture: the host pipeline

`src/FormatDocxHost/Program.cs` is a single-file console host. Its flow (every step matters for correct invocation):

1. **Parse args** → `input` (first positional, `.doc`/`.docx` only), `--rules`, `--output`. Writes **UTF-8 with no BOM** to both streams.
2. **Load rule** (`LoadRule`) — accepts two JSON shapes: a single `FormatRule` object, **or** the plugin wrapper `{ "Rules": [...], "ActiveRuleIndex": n }` (case-insensitive keys); the wrapper is unwrapped by index. Then **`ResolveTableFontsFromBody`** fills empty `Table.FontCN/FontEN` from the body style (宋体 → Times New Roman for Western) — this host-layer policy (not in the engine) realizes the skill's "table font follows body" default.
3. **Spawn a hidden Word COM instance** with `DisplayAlerts=none` and `AutomationSecurity=ForceDisable` (so opening an untrusted doc cannot run macros).
4. **Open the doc**, then **guard against protected view / read-only** before formatting. Returns **exit code 3** if either condition holds — this maps to user guidance: close the Word window that has the doc open, or click "Enable Editing" on the protected-view banner.
5. **Call `Formatter.FormatDocument(doc, rule, progressCb)`** (from `WordFormatter.dll`) — this is where all actual formatting happens. Progress is forwarded to **stderr**, one line per phase.
6. **Save** — `SaveAs2(output)` if `--output` differs from input, else in-place `Save()`.
7. **Emit a single-line JSON summary to stdout** (see below), close/quit Word, release COM objects.

### Process contract (essential when invoking the skill)
- **stdout**: on success, exactly **one line** of JSON, e.g.
  `{"ok":true,"file":"...","rule":"水利工程报告","paragraphs":27,"headings":7,"body":6,"images":0,"figCaptions":1,"tblCaptions":1,"extractedTableTitles":0,"skipped":12,"elapsedMs":602}`.
- **stderr**: progress lines prefixed `[FormatDocx]`, errors prefixed `[FormatDocx][ERROR]`.
- **Exit codes**: `0` ok · `1` formatting error · `2` arg / rule-file error · `3` doc locked / read-only / protected view. Parse stdout for stats, read stderr for diagnostics, and translate code 3 into actionable user steps.

## Rules & rule customization

The schema (`references/rules-schema.md`) and a complete example (`references/default-rule.json`) define a `FormatRule`: `PageSetup`, `Styles` (fixed keys `heading1`–`heading4`, `body`, `figCaption`, `tblCaption`), `Table`, `Image`, `SkipPages`. To customize, copy `default-rule.json`, edit, pass via `--rules`.

Heading/caption **recognition is engine-internal and not rule-configurable**: `第X章/篇/部`→h1, `第X节`→h2, numeric (`1`, `1.1`, `1.1.1`, `1.1.1.1`)→h1–h4 by depth (incl. Word auto-numbering), existing `标题 N` style/outline → by level. Hard limits: >60 chars never a heading; `图/表…` paragraphs >40 chars are treated as body. Table paragraphs are excluded from paragraph formatting (governed only by `Table.*`). `Table.ExtractTitles=true` lifts a title out of the table's first merged row into a standalone caption paragraph.

## Gotchas

- **Engine source is elsewhere** — *except* table-font resolution. To change *how* paragraphs are classified, headings detected, or styles applied, edit `D:\Claude Code\WordTools\src\WordFormatter\`, rebuild it, and refresh `scripts/WordFormatter.dll`. **One exception lives in this repo's `Program.cs`: `ResolveTableFontsFromBody`** fills empty table fonts from the body style before calling the engine — so the "table font follows body" default is a host-layer policy that keeps the shared engine untouched. Change table-font-default behavior *there*, not in WordTools.
- **ProtectedViewWindows.Count is unreliable** (reports a constant 1 on some Word versions). The host correctly enumerates `ProtectedViewWindows` and matches by full path — `Program.cs:IsInProtectedView`. Don't "simplify" it back to `Count`.
- **Platform**: Windows + desktop Microsoft Word 2010+. WPS and Word for the web are unsupported. The COM call must not be interrupted by the user opening the same doc in another Word window (→ read-only → exit 3).
- **Sibling skill `design-report-review`** reviews an already-formatted report. They pair as format → review; see its own SKILL.md.
