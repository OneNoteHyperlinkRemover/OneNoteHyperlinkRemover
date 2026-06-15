using System;
using System.Threading;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// Monitors clipboard on a background thread and removes zero-width spaces.
    /// Uses System.Threading.Timer to avoid interfering with OneNote's UI thread.
    /// </summary>
    internal sealed class ClipboardMonitor : IDisposable
    {
        private const string ZeroWidthSpace = "​";
        private readonly Timer _timer;
        private string _lastText = "";
        private bool _disposed;

        public ClipboardMonitor()
        {
            _timer = new Timer(OnTick, null, 500, 500);
            Log("ClipboardMonitor started (threading timer)");
        }

        private void OnTick(object state)
        {
            if (_disposed) return;

            try
            {
                // Must use STA thread for clipboard access
                string text = null;
                var thread = new Thread(() =>
                {
                    try
                    {
                        if (System.Windows.Forms.Clipboard.ContainsText())
                            text = System.Windows.Forms.Clipboard.GetText();
                    }
                    catch { }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join(200);

                if (string.IsNullOrEmpty(text) || text == _lastText) return;
                _lastText = text;

                if (!text.Contains(ZeroWidthSpace)) return;

                string cleaned = text.Replace(ZeroWidthSpace, "");
                if (cleaned == text) return;

                // Write back on STA thread
                var writeThread = new Thread(() =>
                {
                    try
                    {
                        System.Windows.Forms.Clipboard.SetText(cleaned);
                    }
                    catch { }
                });
                writeThread.SetApartmentState(ApartmentState.STA);
                writeThread.Start();
                writeThread.Join(200);

                _lastText = cleaned;
                Log("RESTORED: " + cleaned.Substring(0, Math.Min(80, cleaned.Length)));
            }
            catch (Exception ex)
            {
                Log("Error: " + ex.Message);
            }
        }

        private static void Log(string msg)
        {
            try
            {
                var path = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OneNoteHyperlinkRemover", "addin.log");
                System.IO.File.AppendAllText(path,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [Clipboard] {msg}{Environment.NewLine}");
            }
            catch { }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                _timer.Dispose();
            }
        }
    }
}
