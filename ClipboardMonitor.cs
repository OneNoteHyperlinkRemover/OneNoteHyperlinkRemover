using System;
using System.Threading;
using System.Windows.Forms;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// Monitors clipboard on a dedicated STA thread and removes zero-width spaces.
    /// Uses a single persistent STA thread to avoid creating threads on every tick,
    /// ensuring clean shutdown when OneNote exits.
    /// </summary>
    internal sealed class ClipboardMonitor : IDisposable
    {
        private const string ZeroWidthSpace = "​";
        private string _lastText = "";
        private volatile bool _disposed;
        private readonly Thread _staThread;
        private readonly AutoResetEvent _wake = new(false);

        public ClipboardMonitor()
        {
            _staThread = new Thread(MonitorLoop)
            {
                IsBackground = true,
                Name = "ClipboardMonitor"
            };
            _staThread.SetApartmentState(ApartmentState.STA);
            _staThread.Start();
            Log("ClipboardMonitor started");
        }

        private void MonitorLoop()
        {
            while (!_disposed)
            {
                try
                {
                    if (Clipboard.ContainsText())
                    {
                        string text = Clipboard.GetText();
                        if (!string.IsNullOrEmpty(text) && text != _lastText)
                        {
                            _lastText = text;
                            if (text.Contains(ZeroWidthSpace))
                            {
                                string cleaned = text.Replace(ZeroWidthSpace, "");
                                if (cleaned != text)
                                {
                                    Clipboard.SetText(cleaned);
                                    _lastText = cleaned;
                                    Log("RESTORED: " + cleaned.Substring(0, Math.Min(80, cleaned.Length)));
                                }
                            }
                        }
                    }
                }
                catch { }

                // Wait 500ms or until disposed
                _wake.WaitOne(500);
            }
            Log("ClipboardMonitor stopped");
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
                _wake.Set(); // Wake up the thread so it can exit
                _staThread.Join(1000); // Wait up to 1 second for clean exit
                _wake.Dispose();
            }
        }
    }
}
