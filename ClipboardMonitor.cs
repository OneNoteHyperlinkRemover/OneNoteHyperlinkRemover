using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// Monitors clipboard changes using SetWinEventHook with EVENT_CLIPBOARD_UPDATE.
    /// This approach uses a callback function instead of window messages,
    /// bypassing the message loop issue in COM add-ins.
    /// Requires Windows 8+.
    /// </summary>
    internal sealed class ClipboardMonitor : IDisposable
    {
        private const uint EVENT_CLIPBOARD_UPDATE = 0x0701;
        private const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        private const string ZeroWidthSpace = "​";

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(
            uint eventMin, uint eventMax,
            IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc,
            uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        private delegate void WinEventDelegate(
            IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild,
            uint dwEventThread, uint dwmsEventTime);

        private IntPtr _hook;
        private WinEventDelegate _delegateRef; // prevent GC from collecting delegate
        private string _lastText = "";
        private bool _disposed;
        private bool _updating; // re-entrancy guard

        public ClipboardMonitor()
        {
            _delegateRef = OnWinEvent;
            _hook = SetWinEventHook(
                EVENT_CLIPBOARD_UPDATE, EVENT_CLIPBOARD_UPDATE,
                IntPtr.Zero, _delegateRef,
                0, 0, WINEVENT_OUTOFCONTEXT);

            if (_hook == IntPtr.Zero)
                Log("SetWinEventHook failed!");
            else
                Log("ClipboardMonitor started (SetWinEventHook)");
        }

        private void OnWinEvent(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild,
            uint dwEventThread, uint dwmsEventTime)
        {
            if (eventType != EVENT_CLIPBOARD_UPDATE || _disposed || _updating)
                return;

            try
            {
                if (!Clipboard.ContainsText()) return;
                string text = Clipboard.GetText();
                if (string.IsNullOrEmpty(text) || text == _lastText) return;
                _lastText = text;

                if (!text.Contains(ZeroWidthSpace)) return;
                string cleaned = text.Replace(ZeroWidthSpace, "");
                if (cleaned == text) return;

                _updating = true;
                try
                {
                    Clipboard.SetText(cleaned);
                }
                finally
                {
                    _updating = false;
                }

                _lastText = cleaned;
                Log("CLEANED: " + cleaned.Substring(0, Math.Min(80, cleaned.Length)));
            }
            catch (Exception ex) { Log("Error: " + ex.Message); }
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
                if (_hook != IntPtr.Zero)
                {
                    UnhookWinEvent(_hook);
                    _hook = IntPtr.Zero;
                }
                Log("ClipboardMonitor stopped");
            }
        }
    }
}
