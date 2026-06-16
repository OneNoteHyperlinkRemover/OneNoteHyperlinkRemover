using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteHyperlinkRemover
{
    internal static class HyperlinkRemover
    {
        // Match <a> tag with :// in href
        // Handles: <a\nhref="...">, nested <span> tags with lang attributes
        // Groups: 1=href, 2=full inner content (including spans)
        private static readonly Regex LinkPattern = new(
            @"<a[\s]+href=""([^""]*://[^""]*)"">(.*?)</a>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

        private static readonly XNamespace Ns =
            XNamespace.Get(OneNoteHelper.OneNoteNamespace);

        public const string ZeroWidthSpace = "​";

        /// <summary>
        /// Remove all auto-converted hyperlinks from the current page.
        /// </summary>
        public static int RemoveHyperlinksFromCurrentPage(OneNoteHelper oneNote)
        {
            var (pageXml, pageId) = GetPageAndId(oneNote);
            if (pageXml == null) return 0;

            var analysis = AnalyzeHyperlinks(pageXml);
            foreach (var link in analysis.Links)
                Logger.Log($"  Link: href=[{link.Href}] text=[{link.DisplayText}] auto={link.IsAutoConverted}");

            if (analysis.AutoConvertedCount == 0)
                return 0;

            return UpdatePage(oneNote, pageId, pageXml);
        }

        /// <summary>
        /// Remove hyperlinks from the current selection only.
        /// </summary>
        public static int RemoveHyperlinksFromSelection(OneNoteHelper oneNote)
        {
            var (pageXml, pageId) = GetPageAndId(oneNote);
            if (pageXml == null) return 0;

            var doc = XDocument.Parse(pageXml);
            var root = doc.Root;
            if (root == null) return 0;

            var ns = root.GetNamespaceOfPrefix(OneNoteHelper.OneNoteNamespacePrefix) ?? Ns;
            int total = 0;

            // Find T elements with selected="all" or inside OE with selected="partial"
            var selectedTs = root.Descendants(ns + "T")
                .Where(t => t.Attribute("selected")?.Value == "all")
                .Concat(root.Descendants(ns + "OE")
                    .Where(oe => oe.Attribute("selected")?.Value == "partial")
                    .SelectMany(oe => oe.Descendants(ns + "T")))
                .Distinct().ToList();

            Logger.Log("Selected T elements: " + selectedTs.Count);

            foreach (var tElement in selectedTs)
            {
                foreach (var cdata in tElement.Nodes().OfType<XCData>())
                {
                    var (processed, count) = StripHyperlinks(cdata.Value);
                    if (count > 0) { cdata.Value = processed; total += count; }
                }
            }

            Logger.Log("Total hyperlinks in selection: " + total);

            if (total > 0)
                SavePage(oneNote, pageId, doc.ToString(SaveOptions.DisableFormatting));

            return total;
        }

        /// <summary>
        /// Strip hyperlinks from a CDATA string. Returns processed text and count.
        /// </summary>
        public static (string Processed, int Count) StripHyperlinks(string cdata)
        {
            int count = 0;
            string result = LinkPattern.Replace(cdata, match =>
            {
                count++;
                string innerContent = match.Groups[2].Value;
                // Strip all HTML tags to get plain text
                string plainText = Regex.Replace(innerContent, "<[^>]+>", "");
                return BreakUrlPattern(plainText);
            });
            return (result, count);
        }

        /// <summary>
        /// Get selected text from the current page.
        /// </summary>
        public static string GetSelectedText(OneNoteHelper oneNote)
        {
            var (pageXml, _) = GetPageAndId(oneNote);
            if (pageXml == null) return null;

            var doc = XDocument.Parse(pageXml);
            var root = doc.Root;
            if (root == null) return null;

            var ns = root.GetNamespaceOfPrefix(OneNoteHelper.OneNoteNamespacePrefix) ?? Ns;

            var selectedTs = root.Descendants(ns + "T")
                .Where(t => t.Attribute("selected")?.Value == "all")
                .Concat(root.Descendants(ns + "OE")
                    .Where(oe => oe.Attribute("selected")?.Value == "partial")
                    .SelectMany(oe => oe.Descendants(ns + "T")))
                .Distinct().ToList();

            var parts = new List<string>();
            foreach (var t in selectedTs)
            {
                foreach (var cdata in t.Nodes().OfType<XCData>())
                {
                    string text = StripHtmlTags(cdata.Value);
                    if (!string.IsNullOrWhiteSpace(text))
                        parts.Add(text);
                }
            }
            return parts.Count > 0 ? string.Join(Environment.NewLine, parts) : null;
        }

        private static string StripHtmlTags(string html)
        {
            return Regex.Replace(html, "<[^>]+>", "");
        }

        /// <summary>
        /// Analyze hyperlinks in page XML.
        /// </summary>
        public static HyperlinkAnalysis AnalyzeHyperlinks(string pageXml)
        {
            var analysis = new HyperlinkAnalysis();
            foreach (Match match in LinkPattern.Matches(pageXml))
            {
                string href = match.Groups[1].Value;
                string innerContent = match.Groups[2].Value;
                string text = Regex.Replace(innerContent, "<[^>]+>", "");
                bool auto = href.Contains("://") && (
                    href == text || text.StartsWith("http") ||
                    text.StartsWith("www.") || text.StartsWith("ftp") ||
                    href.EndsWith(text));
                analysis.Links.Add(new HyperlinkInfo { Href = href, DisplayText = text, IsAutoConverted = auto });
            }
            return analysis;
        }

        #region Helpers

        private static (string PageXml, string PageId) GetPageAndId(OneNoteHelper oneNote)
        {
            string pageXml = oneNote.GetCurrentPageXml();
            string pageId = OneNoteHelper.ExtractPageId(pageXml);
            if (string.IsNullOrEmpty(pageId))
            {
                Logger.Log("ExtractPageId returned null/empty");
                return (null, null);
            }
            return (pageXml, pageId);
        }

        private static int UpdatePage(OneNoteHelper oneNote, string pageId, string pageXml)
        {
            var (modified, count) = StripHyperlinks(pageXml);
            Logger.Log("RemovedCount=" + count);

            if (count > 0) SavePage(oneNote, pageId, modified);
            return count;
        }

        private static void SavePage(OneNoteHelper oneNote, string pageId, string xml)
        {
            Logger.Log("Calling UpdatePageContent...");
            try
            {
                oneNote.UpdatePageContent(pageId, xml);
                Logger.Log("UpdatePageContent succeeded");
            }
            catch (Exception ex) { Logger.Log("UpdatePageContent error: " + ex); }
        }

        /// <summary>
        /// Insert zero-width spaces at :// and www. to break URL pattern.
        /// </summary>
        private static string BreakUrlPattern(string url)
        {
            return url.Replace("://", "://​").Replace("www.", "www.​");
        }

        #endregion
    }

    internal class HyperlinkAnalysis
    {
        public List<HyperlinkInfo> Links { get; } = new();
        public int AutoConvertedCount => Links.Count(l => l.IsAutoConverted);
    }

    internal class HyperlinkInfo
    {
        public string Href { get; set; }
        public string DisplayText { get; set; }
        public bool IsAutoConverted { get; set; }
    }
}
