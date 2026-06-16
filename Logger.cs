using System;
using System.Diagnostics;
using System.IO;

namespace OneNoteHyperlinkRemover
{
    internal static class Logger
    {
        private static readonly string LogDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OneNoteHyperlinkRemover");
        private static readonly string LogFile = Path.Combine(LogDir, "addin.log");

        [Conditional("DEBUG")]
        public static void Log(string message)
        {
            try
            {
                if (!Directory.Exists(LogDir))
                    Directory.CreateDirectory(LogDir);
                File.AppendAllText(LogFile,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}");
            }
            catch { }
        }
    }
}
