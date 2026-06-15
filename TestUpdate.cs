using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.OneNote;

class TestUpdate
{
    static void Main()
    {
        Console.WriteLine("Creating OneNote Application...");
        var app = new Application();

        try
        {
            // Get current page
            Console.WriteLine("Getting current page...");
            var windows = app.Windows;
            var window = windows.CurrentWindow;
            string pageId = window.CurrentPageId;
            Console.WriteLine("PageId: " + pageId);

            Marshal.ReleaseComObject(window);
            Marshal.ReleaseComObject(windows);

            // Read page XML
            string pageXml;
            app.GetPageContent(pageId, out pageXml, PageInfo.piAll);
            Console.WriteLine("Page XML length: " + pageXml.Length);

            // Check for hyperlinks
            int hrefCount = CountOccurrences(pageXml, "href=\"http");
            Console.WriteLine("Hyperlinks found: " + hrefCount);

            if (hrefCount == 0)
            {
                Console.WriteLine("No hyperlinks to remove.");
                return;
            }

            // Show original CDATA
            int cdataIdx = pageXml.IndexOf("CDATA[<a");
            if (cdataIdx >= 0)
            {
                Console.WriteLine("Original CDATA (200 chars): " + pageXml.Substring(cdataIdx, Math.Min(200, pageXml.Length - cdataIdx)));
            }

            // Remove hyperlinks using string replacement
            string modified = System.Text.RegularExpressions.Regex.Replace(
                pageXml,
                @"<a\s+href=""(https?://[^""]*)"">\1</a>",
                "$1",
                System.Text.RegularExpressions.RegexOptions.Singleline |
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            int newHrefCount = CountOccurrences(modified, "href=\"http");
            Console.WriteLine("Hyperlinks after removal: " + newHrefCount);

            // Show modified CDATA
            cdataIdx = modified.IndexOf("CDATA[http");
            if (cdataIdx >= 0)
            {
                Console.WriteLine("Modified CDATA (200 chars): " + modified.Substring(cdataIdx, Math.Min(200, modified.Length - cdataIdx)));
            }

            // Try UpdatePageContent
            Console.WriteLine("Calling UpdatePageContent...");
            try
            {
                app.UpdatePageContent(modified, DateTime.MinValue, XMLSchema.xs2013, true);
                Console.WriteLine("UpdatePageContent returned successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("UpdatePageContent error: " + ex.Message);
                return;
            }

            // Verify
            Console.WriteLine("Re-reading page...");
            string verifyXml;
            app.GetPageContent(pageId, out verifyXml, PageInfo.piAll);
            int verifyCount = CountOccurrences(verifyXml, "href=\"http");
            Console.WriteLine("Hyperlinks after update: " + verifyCount);

            if (verifyCount < hrefCount)
            {
                Console.WriteLine("SUCCESS: Hyperlinks were removed!");
            }
            else
            {
                Console.WriteLine("FAILED: Hyperlinks still present after update.");
                Console.WriteLine("Page XML unchanged: " + (verifyXml == pageXml ? "YES (identical)" : "NO (different)"));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex);
        }
        finally
        {
            Marshal.ReleaseComObject(app);
        }

        Console.WriteLine("Press Enter to exit...");
        Console.ReadLine();
    }

    static int CountOccurrences(string text, string pattern)
    {
        int count = 0, idx = 0;
        while ((idx = text.IndexOf(pattern, idx, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            count++;
            idx += pattern.Length;
        }
        return count;
    }
}
