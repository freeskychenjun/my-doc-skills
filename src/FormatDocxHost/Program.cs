using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordFormatter.Engine;
using WordFormatter.Models;

namespace FormatDocxHost
{
    /// <summary>
    /// Word 设计报告一键排版的命令行宿主。
    /// 通过 COM 自动化启动独立 Word 实例，复用 WordFormatter 引擎完成排版。
    ///
    /// 用法:
    ///   FormatDocx.exe &lt;input.docx&gt; [--rules rules.json] [--output out.docx]
    ///
    /// 退出码:
    ///   0 成功（stdout 输出一行 JSON 统计）
    ///   1 排版过程出错
    ///   2 参数/规则文件错误
    ///   3 文档被占用、只读或处于受保护视图
    /// </summary>
    internal static class Program
    {
        private static readonly TextWriter Err =
            new StreamWriter(Console.OpenStandardError(), new UTF8Encoding(false)) { AutoFlush = true };

        [STAThread]
        static int Main(string[] args)
        {
            Console.OutputEncoding = new UTF8Encoding(false);

            string input = null, rulesPath = null, output = null;
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--rules":
                        if (++i >= args.Length) return ArgError("--rules 缺少路径");
                        rulesPath = args[i];
                        break;
                    case "--output":
                        if (++i >= args.Length) return ArgError("--output 缺少路径");
                        output = args[i];
                        break;
                    case "-h":
                    case "--help":
                        PrintUsage();
                        return 0;
                    default:
                        if (input == null) input = args[i];
                        else return ArgError("未知参数: " + args[i]);
                        break;
                }
            }

            if (input == null)
            {
                PrintUsage();
                return 2;
            }
            input = Path.GetFullPath(input);
            if (!File.Exists(input)) return ArgError("文件不存在: " + input);
            if (!input.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                && !input.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                return ArgError("仅支持 .doc/.docx 文件: " + input);
            if (output != null) output = Path.GetFullPath(output);

            FormatRule rule;
            try
            {
                rule = LoadRule(rulesPath);
            }
            catch (Exception ex)
            {
                return ArgError("规则加载失败: " + ex.Message);
            }

            // 表格字体默认跟随正文：规则未显式指定表格中西文字体时，用正文字体补齐。
            // 在宿主层解析（不改动共享引擎行为），字号保持不动（引擎 FontSize=0 即保留原字号）。
            ResolveTableFontsFromBody(rule);

            Word.Application app = null;
            Word.Document doc = null;
            int exitCode = 1;
            try
            {
                app = new Word.Application();
                app.Visible = false;
                app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                // 禁用宏自动运行，避免打开陌生文档时触发宏
                try { app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable; } catch { }

                Log("打开文档: " + input);
                doc = app.Documents.Open(
                    FileName: input,
                    ConfirmConversions: false,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);

                // ProtectedViewWindows.Count 在部分 Word 版本上恒为 1（已知误报），
                // 必须枚举窗口逐个比对文档路径来判定。
                if (IsInProtectedView(app, input))
                {
                    Error("文档在「受保护的视图」中打开，无法排版。请先在 Word 中打开该文件并点击「启用编辑」。");
                    exitCode = 3;
                    return exitCode;
                }
                if (doc.ReadOnly)
                {
                    Error("文档以只读方式打开（可能被其他 Word 窗口占用或文件属性为只读），请关闭后重试。");
                    exitCode = 3;
                    return exitCode;
                }

                Log("开始排版，规则: " + (rule.Name ?? "默认"));
                FormatResult result = Formatter.FormatDocument(doc, rule, p =>
                {
                    if (p != null && !string.IsNullOrEmpty(p.Detail))
                        Log(string.Format("{0} {1}/{2} {3}", p.Phase, p.Current, p.Total, p.Detail));
                }).GetAwaiter().GetResult();

                string savedTo;
                if (output != null && !StringComparer.OrdinalIgnoreCase.Equals(output, input))
                {
                    string outDir = Path.GetDirectoryName(output);
                    if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                        Directory.CreateDirectory(outDir);
                    doc.SaveAs2(FileName: output);
                    savedTo = output;
                }
                else
                {
                    doc.Save();
                    savedTo = input;
                }

                var summary = JsonConvert.SerializeObject(new
                {
                    ok = true,
                    file = savedTo,
                    rule = rule.Name,
                    paragraphs = result.Total,
                    headings = result.Headings,
                    body = result.Body,
                    images = result.Images,
                    figCaptions = result.FigCaptions,
                    tblCaptions = result.TblCaptions,
                    extractedTableTitles = result.ExtractedTitles,
                    skipped = result.Skipped,
                    elapsedMs = Math.Round(result.TotalMs)
                });
                Console.WriteLine(summary);
                exitCode = 0;
                return exitCode;
            }
            catch (Exception ex)
            {
                Error("排版失败: " + ex.Message);
                return exitCode;
            }
            finally
            {
                try { if (doc != null) doc.Close(SaveChanges: false); } catch { }
                try { if (app != null) app.Quit(SaveChanges: false); } catch { }
                try { if (doc != null) Marshal.FinalReleaseComObject(doc); } catch { }
                try { if (app != null) Marshal.FinalReleaseComObject(app); } catch { }
            }
        }

        /// <summary>
        /// 规则加载优先级: --rules 指定文件 &gt; %AppData%\CJTools\rules.json &gt; 内置默认规则。
        /// 支持两种 JSON 形态: 单个 FormatRule 对象，或插件导出的 { Rules: [...], ActiveRuleIndex: n } 包装。
        /// </summary>
        private static FormatRule LoadRule(string rulesPath)
        {
            string path = rulesPath;
            if (path == null)
            {
                var appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "CJTools", "rules.json");
                if (File.Exists(appDataPath)) path = appDataPath;
            }
            else if (!File.Exists(path))
            {
                throw new FileNotFoundException("规则文件不存在: " + path);
            }

            if (path == null)
            {
                Log("未找到规则文件，使用内置默认规则（水利工程报告）");
                return FormatRule.CreateDefault();
            }

            Log("加载规则: " + path);
            var json = JToken.Parse(File.ReadAllText(path, Encoding.UTF8));
            var rules = json["Rules"] ?? json["rules"];
            if (rules != null && rules.Type == JTokenType.Array && rules.Any())
            {
                int idx = (int?)(json["ActiveRuleIndex"] ?? json["activeRuleIndex"]) ?? 0;
                idx = Math.Max(0, Math.Min(idx, rules.Count() - 1));
                return rules[idx].ToObject<FormatRule>();
            }
            return json.ToObject<FormatRule>();
        }

        /// <summary>
        /// 表格字体默认跟随正文：当规则未显式指定表格中西文字体（留空）时，用正文字体补齐。
        ///   表格中文字体留空 → 取正文中文字体
        ///   表格西文字体留空 → 正文为宋体时用 Times New Roman，否则取正文西文字体
        ///   字号不动（保留表格原字号，引擎 FontSize=0 即不设置）
        /// 用户显式指定的表格字体（非空）优先，不会被覆盖。
        /// 仅在本 skill 宿主层解析，不改动共享引擎的「留空 = 保留原字体」语义。
        /// </summary>
        private static void ResolveTableFontsFromBody(FormatRule rule)
        {
            var table = rule?.Table;
            if (table == null) return;

            bool needCN = string.IsNullOrEmpty(table.FontCN);
            bool needEN = string.IsNullOrEmpty(table.FontEN);
            if (!needCN && !needEN) return;

            StyleRule body = null;
            rule.Styles?.TryGetValue("body", out body);
            if (body == null) return;

            string resolvedCN = table.FontCN;
            string resolvedEN = table.FontEN;

            if (needCN && !string.IsNullOrEmpty(body.FontCN))
                resolvedCN = body.FontCN;

            if (needEN)
            {
                // 正文为宋体时，表格西文字体默认 Times New Roman；否则跟随正文西文字体
                resolvedEN = (body.FontCN == "宋体")
                    ? "Times New Roman"
                    : body.FontEN;
            }

            if (resolvedCN != table.FontCN || resolvedEN != table.FontEN)
            {
                table.FontCN = resolvedCN;
                table.FontEN = resolvedEN;
                Log("表格字体未指定，跟随正文：" + (resolvedCN ?? "（保留原字体）")
                    + " / " + (resolvedEN ?? "（保留原字体）"));
            }
        }

        /// <summary>
        /// 枚举 ProtectedViewWindows，判断指定文件是否在受保护视图中打开。
        /// 不能用 Count（部分版本恒为 1）。
        /// </summary>
        private static bool IsInProtectedView(Word.Application app, string fullPath)
        {
            try
            {
                foreach (Word.ProtectedViewWindow pvw in app.ProtectedViewWindows)
                {
                    try
                    {
                        var pvDoc = pvw.Document;
                        if (pvDoc != null && string.Equals(
                                pvDoc.FullName, fullPath, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    catch { /* 单个窗口访问失败则跳过 */ }
                }
            }
            catch { /* 枚举失败视为不在受保护视图 */ }
            return false;
        }

        private static void PrintUsage()
        {
            Err.WriteLine(
                "用法: FormatDocx.exe <input.docx> [--rules rules.json] [--output out.docx]\n" +
                "  --rules   排版规则 JSON（FormatRule 或插件导出的 RulesData 包装），\n" +
                "            缺省依次尝试 %AppData%\\CJTools\\rules.json、内置默认规则\n" +
                "  --output  另存到指定路径；缺省为原地保存\n" +
                "成功时 stdout 输出一行 JSON 统计，进度信息输出到 stderr。");
        }

        private static int ArgError(string msg)
        {
            Error(msg);
            return 2;
        }

        private static void Log(string msg)
        {
            Err.WriteLine("[FormatDocx] " + msg);
        }

        private static void Error(string msg)
        {
            Err.WriteLine("[FormatDocx][ERROR] " + msg);
        }
    }
}
