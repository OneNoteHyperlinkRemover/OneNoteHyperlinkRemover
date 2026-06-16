using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// Monitors clipboard changes by polling GetClipboardSequenceNumber.
    /// This is the only approach that works in OneNote COM add-ins because
    /// OneNote's message loop does not pump messages for windows/hooks
    /// created by add-ins (AddClipboardFormatListener and SetWinEventHook
    /// both require a message pump to deliver callbacks).
    /// </summary>
    internal sealed class ClipboardMonitor : IDisposable
    {
        private const string ZeroWidthSpace = "​";

        [DllImport("user32.dll")]
        private static extern uint GetClipboardSequenceNumber();

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        private readonly Thread _thread;
        private readonly ManualResetEvent _stop = new(false);
        private readonly int _interval;
        private uint _lastSeq;
        private string _lastText = "";
        private bool _disposed;

        public ClipboardMonitor(int intervalMs = 300)
        {
            _interval = Math.Max(100, intervalMs);
            _lastSeq = GetClipboardSequenceNumber();
            _thread = new Thread(MonitorLoop)
            {
                IsBackground = true,
                Name = "ClipboardMonitor"
            };
            _thread.Start();
            Log("ClipboardMonitor started (polling " + _interval + "ms)");
        }

        private void MonitorLoop()
        {
            while (!_stop.WaitOne(_interval))
            {
                try
                {
                    uint seq = GetClipboardSequenceNumber();
                    if (seq == _lastSeq) continue;
                    _lastSeq = seq;

                    string text = null;
                    var readThread = new Thread(() =>
                    {
                        try
                        {
                            if (Clipboard.ContainsText())
                                text = Clipboard.GetText();
                        }
                        catch { }
                    });
                    readThread.SetApartmentState(ApartmentState.STA);
                    readThread.Start();
                    readThread.Join(500);

                    if (string.IsNullOrEmpty(text) || text == _lastText) continue;
                    _lastText = text;

                    if (!text.Contains(ZeroWidthSpace)) continue;
                    if (!IsOneNoteForeground()) continue;
                    string cleaned = text.Replace(ZeroWidthSpace, "");
                    if (cleaned == text) continue;

                    var writeThread = new Thread(() =>
                    {
                        try { Clipboard.SetText(cleaned); }
                        catch { }
                    });
                    writeThread.SetApartmentState(ApartmentState.STA);
                    writeThread.Start();
                    writeThread.Join(500);

                    _lastText = cleaned;
                    Log("CLEANED: " + cleaned.Substring(0, Math.Min(80, cleaned.Length)));
                }
                catch { }
            }
        }

        private static bool IsOneNoteForeground()
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == IntPtr.Zero) return false;
                GetWindowThreadProcessId(hwnd, out uint pid);
                var proc = Process.GetProcessById((int)pid);
                return proc.ProcessName.Equals("ONENOTE", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        [Conditional("DEBUG")]
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
                _stop.Set();
                _thread.Join(1000);
                _stop.Dispose();
                Log("ClipboardMonitor stopped");
            }
        }
    }
}
