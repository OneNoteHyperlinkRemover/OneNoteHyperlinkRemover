using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OneNoteHyperlinkRemover
{
    internal static class HyperlinkRemover
    {
        // Match any <a> tag where href contains :// (any protocol)
        // Group 1: full href, Group 2: display text
        private static readonly Regex AnyHyperlinkPattern = new(
            @"<a\s+href=""([^""]*://[^""]*)"">([^<]*)</a>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

        // Track removed URLs per page to avoid infinite loop
        // (OneNote re-converts URLs after UpdatePageContent)
        private static readonly Dictionary<string, HashSet<string>> _removedUrls = new();

        private static void Log(string msg)
        {
            try
            {
                var path = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OneNoteHyperlinkRemover", "addin.log");
                System.IO.File.AppendAllText(path,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [HyperlinkRemover] {msg}{Environment.NewLine}");
            }
            catch { }
        }

        public static int RemoveHyperlinksFromCurrentPage(OneNoteHelper oneNote)
        {
            string pageXml = oneNote.GetCurrentPageXml();
            string pageId = OneNoteHelper.ExtractPageId(pageXml);

            if (string.IsNullOrEmpty(pageId))
            {
                Log("ExtractPageId returned null/empty");
                return 0;
            }

            // Find all auto-converted URLs on the page
            var analysis = AnalyzeHyperlinks(pageXml);

            // Log all found hyperlinks for debugging
            foreach (var link in analysis.Links)
            {
                Log($"  Link: href=[{link.Href}] text=[{link.DisplayText}] auto={link.IsAutoConverted}");
            }

            if (analysis.AutoConvertedCount == 0)
            {
                // No auto-converted hyperlinks, clear tracking for this page
                _removedUrls.Remove(pageId);
                return 0;
            }

            // Get or create the set of already-removed URLs for this page
            if (!_removedUrls.TryGetValue(pageId, out var removed))
            {
                removed = new HashSet<string>();
                _removedUrls[pageId] = removed;
            }

            // Check if there are NEW URLs that we haven't processed yet
            var newUrls = analysis.Links
                .Where(l => l.IsAutoConverted && !removed.Contains(l.Href))
                .ToList();

            if (newUrls.Count == 0)
            {
                // All URLs already processed, skip update to avoid infinite loop
                return 0;
            }

            Log("PageId=" + pageId + ", newUrls=" + newUrls.Count);

            // Mark these URLs as removed
            foreach (var url in newUrls)
                removed.Add(url.Href);

            // Remove hyperlinks
            var result = RemoveHyperlinksFromXml(pageXml);
            Log("RemovedCount=" + result.RemovedCount);

            // Debug: save modified XML to verify zero-width space
            try
            {
                var debugDir = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OneNoteHyperlinkRemover");
                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(debugDir, "page_after.xml"), result.ModifiedXml);
            }
            catch { }

            if (result.RemovedCount > 0)
            {
                // Debug: save modified XML
                try
                {
                    var debugPath = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "OneNoteHyperlinkRemover", "page_after.xml");
                    System.IO.File.WriteAllText(debugPath, result.ModifiedXml);
                    Log("Saved modified XML to " + debugPath);
                }
                catch (Exception ex) { Log("Debug save error: " + ex.Message); }

                Log("Calling UpdatePageContent...");
                try
                {
                    oneNote.UpdatePageContent(pageId, result.ModifiedXml);
                    Log("UpdatePageContent succeeded");
                }
                catch (Exception ex)
                {
                    Log("UpdatePageContent error: " + ex);
                }
            }

            return result.RemovedCount;
        }

        public static (string ModifiedXml, int RemovedCount) RemoveHyperlinksFromXml(string pageXml)
        {
            int count = 0;
            string result = pageXml;

            // Match any <a> tag with :// in href
            // Remove <a> tag and break URL pattern in display text
            result = AnyHyperlinkPattern.Replace(result, match =>
            {
                count++;
                string displayText = match.Groups[2].Value;
                return BreakUrlPattern(displayText);
            });

            return (result, count);
        }

        /// <summary>
        /// Break URL pattern by inserting zero-width spaces at key positions:
        /// after :// and after www. to prevent OneNote auto-conversion.
        /// </summary>
        private static string BreakUrlPattern(string url)
        {
            // Insert after :// (e.g., https://​www.baidu.com)
            string result = url.Replace("://", "://​");
            // Insert after www. (e.g., https://www.​baidu.com)
            result = result.Replace("www.", "www.​");
            return result;
        }

        public static HyperlinkAnalysis AnalyzeHyperlinks(string pageXml)
        {
            var analysis = new HyperlinkAnalysis();
            var pattern = new Regex(@"<a\s+href=""([^""]*)"">([^<]*)</a>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            var matches = pattern.Matches(pageXml);
            foreach (Match match in matches)
            {
                string href = match.Groups[1].Value;
                string displayText = match.Groups[2].Value;
                // Auto-converted if display text looks like a URL
                bool isAutoConverted = href.Contains("://") && (
                    href == displayText ||
                    displayText.StartsWith("http") ||
                    displayText.StartsWith("www.") ||
                    displayText.StartsWith("ftp") ||
                    href.EndsWith(displayText));
                analysis.Links.Add(new HyperlinkInfo
                {
                    Href = href,
                    DisplayText = displayText,
                    IsAutoConverted = isAutoConverted
                });
            }
            return analysis;
        }
    }

    internal class HyperlinkAnalysis
    {
        public List<HyperlinkInfo> Links { get; } = new();
        public int TotalCount => Links.Count;
        public int AutoConvertedCount { get { int c = 0; foreach (var l in Links) if (l.IsAutoConverted) c++; return c; } }
        public int ManualCount { get { int c = 0; foreach (var l in Links) if (!l.IsAutoConverted) c++; return c; } }
    }

    internal class HyperlinkInfo
    {
        public string Href { get; set; }
        public string DisplayText { get; set; }
        public bool IsAutoConverted { get; set; }
    }
}
