using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// 超链接移除核心逻辑。
    /// 工作原理：
    ///   1. 获取当前页面 XML
    ///   2. 解析 XML 中的文本元素（OE/T 节点）
    ///   3. 在 CDATA 内容中查找自动转换的 URL 超链接（&lt;a href="..."&gt;text&lt;/a&gt;）
    ///   4. 将超链接替换为纯文本
    ///   5. 写回页面
    /// </summary>
    internal static class HyperlinkRemover
    {
        // 匹配 CDATA 中的 <a href="...">text</a> 标签
        // Group 1: href 属性值（URL）
        // Group 2: 链接显示文本
        private static readonly Regex HyperlinkPattern = new(
            @"<a\s+href=""([^""]*)"">([^<]*)</a>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // 匹配自动转换的裸 URL 超链接（如 <a href="https://example.com">https://example.com</a>）
        // 这种情况下 href 和显示文本相同，是 OneNote 自动转换的特征
        private static readonly Regex AutoConvertedPattern = new(
            @"<a\s+href=""(https?://[^""]*)"">\1</a>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // 匹配更广泛的自动转换模式（包括显示文本是 URL 的截断形式）
        private static readonly Regex AutoConvertedLoosePattern = new(
            @"<a\s+href=""(https?://[^""]*)"">((?:https?://)?[^<]{1,200})</a>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// 从当前活动页面移除自动转换的超链接。
        /// </summary>
        /// <param name="oneNote">OneNoteHelper 实例</param>
        /// <returns>移除的超链接数量</returns>
        public static int RemoveHyperlinksFromCurrentPage(OneNoteHelper oneNote)
        {
            string pageXml = oneNote.GetCurrentPageXml();
            string pageId = OneNoteHelper.ExtractPageId(pageXml);

            if (string.IsNullOrEmpty(pageId))
                return 0;

            var result = RemoveHyperlinksFromXml(pageXml);

            if (result.RemovedCount > 0)
            {
                oneNote.UpdatePageContent(pageId, result.ModifiedXml);
            }

            return result.RemovedCount;
        }

        /// <summary>
        /// 从页面 XML 中移除自动转换的超链接。
        /// </summary>
        /// <param name="pageXml">原始页面 XML</param>
        /// <returns>修改后的 XML 和移除计数</returns>
        public static (string ModifiedXml, int RemovedCount) RemoveHyperlinksFromXml(string pageXml)
        {
            var doc = XDocument.Parse(pageXml);
            var root = doc.Root;
            if (root == null) return (pageXml, 0);

            var ns = root.GetNamespaceOfPrefix(OneNoteHelper.OneNoteNamespacePrefix)
                     ?? XNamespace.Get(OneNoteHelper.OneNoteNamespace);

            int totalCount = 0;

            // 遍历所有 Outline → OE → T 节点（文本运行）
            // 页面结构: Page > Outline > OEChildren > OE > T (CDATA)
            var outlines = root.Descendants(ns + "Outline");
            foreach (var outline in outlines)
            {
                var textElements = outline.Descendants(ns + "T");
                foreach (var tElement in textElements)
                {
                    int count = ProcessTextElement(tElement);
                    totalCount += count;
                }
            }

            return (doc.ToString(), totalCount);
        }

        /// <summary>
        /// 处理单个 T（文本）元素中的 CDATA 内容。
        /// </summary>
        /// <returns>该元素中移除的超链接数量</returns>
        private static int ProcessTextElement(XElement tElement)
        {
            int totalCount = 0;

            // T 元素包含 CDATA 节点，CDATA 中包含 HTML 标签（如 <a>）
            var cdataNodes = tElement.Nodes().OfType<XCData>().ToList();

            foreach (var cdata in cdataNodes)
            {
                var (processed, count) = StripAutoHyperlinks(cdata.Value);
                if (count > 0)
                {
                    cdata.ReplaceWith(new XCData(processed));
                    totalCount += count;
                }
            }

            return totalCount;
        }

        /// <summary>
        /// 从 CDATA 内容中移除自动转换的 URL 超链接。
        /// 只移除 href 和显示文本相同（或显示文本是 URL）的超链接，
        /// 保留用户手动创建的有意义的超链接文本。
        /// </summary>
        /// <param name="cdataContent">CDATA 原始内容</param>
        /// <returns>处理后的内容和移除数量</returns>
        public static (string Processed, int Count) StripAutoHyperlinks(string cdataContent)
        {
            if (string.IsNullOrEmpty(cdataContent))
                return (cdataContent, 0);

            int count = 0;

            // 第一步：精确匹配 — href 和显示文本完全相同
            // 这是 OneNote 自动转换的最典型特征
            string result = AutoConvertedPattern.Replace(cdataContent, match =>
            {
                count++;
                return match.Groups[1].Value; // 只保留 URL 文本
            });

            // 第二步：宽松匹配 — 显示文本看起来像 URL（以 http/https 开头）
            // 但 href 和显示文本可能不完全相同（OneNote 有时会截断显示文本）
            result = AutoConvertedLoosePattern.Replace(result, match =>
            {
                string href = match.Groups[1].Value;
                string displayText = match.Groups[2].Value;

                // 如果显示文本是 URL 的前缀（被截断），也认为是自动转换
                if (displayText.StartsWith("http") && href.StartsWith(displayText))
                {
                    count++;
                    return href; // 用完整 URL 替换截断的显示文本
                }

                // 否则保留原样（可能是用户手动创建的有意义链接）
                return match.Value;
            });

            return (result, count);
        }

        /// <summary>
        /// 分析页面中的超链接情况（用于调试和统计）。
        /// </summary>
        public static HyperlinkAnalysis AnalyzeHyperlinks(string pageXml)
        {
            var doc = XDocument.Parse(pageXml);
            var root = doc.Root;
            var analysis = new HyperlinkAnalysis();

            if (root == null) return analysis;

            var ns = root.GetNamespaceOfPrefix(OneNoteHelper.OneNoteNamespacePrefix)
                     ?? XNamespace.Get(OneNoteHelper.OneNoteNamespace);

            var textElements = root.Descendants(ns + "T");
            foreach (var tElement in textElements)
            {
                foreach (var cdata in tElement.Nodes().OfType<XCData>())
                {
                    var matches = HyperlinkPattern.Matches(cdata.Value);
                    foreach (Match match in matches)
                    {
                        string href = match.Groups[1].Value;
                        string displayText = match.Groups[2].Value;
                        bool isAutoConverted = href == displayText ||
                                               (displayText.StartsWith("http") && href.StartsWith(displayText));

                        analysis.Links.Add(new HyperlinkInfo
                        {
                            Href = href,
                            DisplayText = displayText,
                            IsAutoConverted = isAutoConverted
                        });
                    }
                }
            }

            return analysis;
        }
    }

    /// <summary>
    /// 超链接分析结果。
    /// </summary>
    internal class HyperlinkAnalysis
    {
        public List<HyperlinkInfo> Links { get; } = new();
        public int TotalCount => Links.Count;
        public int AutoConvertedCount => Links.Count(l => l.IsAutoConverted);
        public int ManualCount => Links.Count(l => !l.IsAutoConverted);
    }

    /// <summary>
    /// 单个超链接信息。
    /// </summary>
    internal class HyperlinkInfo
    {
        public string Href { get; set; }
        public string DisplayText { get; set; }
        public bool IsAutoConverted { get; set; }
    }
}
