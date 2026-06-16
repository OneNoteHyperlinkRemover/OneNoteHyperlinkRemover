using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// Monitors clipboard changes using AddClipboardFormatListener (event-driven, no polling).
    /// Creates a hidden message-only window to receive WM_CLIPBOARDUPDATE messages.
    /// </summary>
    internal sealed class ClipboardMonitor : NativeWindow, IDisposable
    {
        private const int WM_CLIPBOARDUPDATE = 0x031D;
        private const string ZeroWidthSpace = "​";

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        private string _lastText = "";
        private bool _disposed;

        public ClipboardMonitor()
        {
            // Create a message-only window
            CreateHandle(new CreateParams
            {
                Parent = (IntPtr)(-3), // HWND_MESSAGE — message-only window
                ClassName = "ClipboardMonitorWindow"
            });
            AddClipboardFormatListener(Handle);
            Log("ClipboardMonitor started (event-driven)");
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE)
            {
                OnClipboardUpdate();
            }
            base.WndProc(ref m);
        }

        private void OnClipboardUpdate()
        {
            if (_disposed) return;
            try
            {
                if (!Clipboard.ContainsText()) return;
                string text = Clipboard.GetText();
                if (string.IsNullOrEmpty(text) || text == _lastText) return;
                _lastText = text;

                if (!text.Contains(ZeroWidthSpace)) return;
                string cleaned = text.Replace(ZeroWidthSpace, "");
                if (cleaned == text) return;

                Clipboard.SetText(cleaned);
                _lastText = cleaned;
                Log("CLEANED: " + cleaned.Substring(0, Math.Min(80, cleaned.Length)));
            }
            catch { }
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
                RemoveClipboardFormatListener(Handle);
                DestroyHandle();
                Log("ClipboardMonitor stopped");
            }
        }
    }
}
