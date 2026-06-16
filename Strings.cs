using System.Collections.Generic;
using System.Globalization;

namespace OneNoteHyperlinkRemover
{
    internal static class Strings
    {
        private static readonly bool IsChinese =
            CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "zh";

        private static readonly Dictionary<string, string[]> Map = new()
        {
            // key: { zh, en }
            ["GroupName"]              = new[] { "超链接工具", "Hyperlink Tools" },
            ["RemovePage"]             = new[] { "移除本页超链接", "Remove Page Links" },
            ["RemovePageScreen"]       = new[] { "移除当前页面的自动超链接", "Remove auto-converted hyperlinks from the current page" },
            ["RemovePageSuper"]        = new[] { "扫描当前 OneNote 页面，将自动转换的 URL 超链接恢复为纯文本。不会影响手动插入的超链接。", "Scan the current OneNote page and restore auto-converted URL hyperlinks to plain text. Manually inserted hyperlinks are preserved." },
            ["RemoveSelection"]        = new[] { "移除选区超链接", "Remove Selection Links" },
            ["RemoveSelectionScreen"]  = new[] { "仅移除选中文字中的超链接", "Remove hyperlinks from selected text only" },
            ["RemoveSelectionSuper"]   = new[] { "只处理当前选中的文字，不影响页面其他内容。", "Process only the selected text without affecting the rest of the page." },
            ["AutoRemove"]             = new[] { "自动移除当前页超链接", "Auto-remove Page Links" },
            ["AutoRemoveScreen"]       = new[] { "开启/关闭自动监控模式", "Toggle automatic hyperlink removal" },
            ["Interval"]               = new[] { "间隔(ms)", "Interval(ms)" },
            ["IntervalScreen"]         = new[] { "自动扫描间隔（毫秒）", "Auto-scan interval (milliseconds)" },
            ["Clipboard"]              = new[] { "自动清理剪贴板零宽空格", "Auto-clean Clipboard ZWS" },
            ["ClipboardScreen"]        = new[] { "开启/关闭自动清除剪贴板中的零宽空格", "Toggle automatic zero-width space cleanup on clipboard" },
            ["CopyCleanText"]          = new[] { "复制移除零宽空格后的文本", "Copy Text Without ZWS" },
            ["CopyCleanTextScreen"]    = new[] { "复制选中文字到剪贴板（清除零宽空格）", "Copy selected text to clipboard with zero-width spaces removed" },
        };

        public static string Get(string key)
        {
            if (Map.TryGetValue(key, out var pair))
                return IsChinese ? pair[0] : pair[1];
            return key;
        }
    }
}
