using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;

namespace OneNoteHyperlinkRemover
{
    [ComVisible(true)]
    [Guid("b7a3d2e1-4f6c-4a8b-9e1d-3c5f7a9b2d4e")]
    [ProgId("OneNoteHyperlinkRemover.AddIn")]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private bool _autoRemoveEnabled;
        private bool _clipboardEnabled = false;
        private int _clipboardInterval = 300;
        private int _autoRemoveInterval = 2000;
        private System.Threading.Timer _autoRemoveTimer;
        private ClipboardMonitor _clipboardMonitor;

        public AddIn()
        {
            Logger.Log("=== Constructor started ===");
            Logger.Log("Assembly: " + Assembly.GetExecutingAssembly().Location);
            Logger.Log("Runtime: " + Environment.Version);
        }

        #region IDTExtensibility2

        public void OnConnection(object app, ext_ConnectMode mode, object inst, ref Array custom)
            => Logger.Log($"=== OnConnection, mode={mode} ===");

        public void OnDisconnection(ext_DisconnectMode mode, ref Array custom)
        {
            _ribbon = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            Logger.Log("=== OnStartupComplete ===");
            if (_clipboardEnabled) StartClipboardMonitor();
        }

        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonId)
        {
            Logger.Log($"=== GetCustomUI, ribbonId={ribbonId} ===");
            try
            {
                using var stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("OneNoteHyperlinkRemover.Ribbon.Ribbon.xml");
                if (stream != null)
                {
                    using var reader = new StreamReader(stream);
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex) { Logger.Log("GetCustomUI error: " + ex); }
            return string.Empty;
        }

        public void Ribbon_Load(IRibbonUI ribbon) { _ribbon = ribbon; }

        #endregion

        #region Ribbon Callbacks

        public void OnRemoveHyperlinks(IRibbonControl control)
        {
            Logger.Log("OnRemoveHyperlinks");
            Execute(() =>
            {
                using var oneNote = new OneNoteHelper();
                return HyperlinkRemover.RemoveHyperlinksFromCurrentPage(oneNote);
            });
        }

        public void OnRemoveSelectionHyperlinks(IRibbonControl control)
        {
            Logger.Log("OnRemoveSelectionHyperlinks");
            Execute(() =>
            {
                using var oneNote = new OneNoteHelper();
                return HyperlinkRemover.RemoveHyperlinksFromSelection(oneNote);
            });
        }

        public void OnToggleAutoRemove(IRibbonControl control, bool pressed)
        {
            Logger.Log("OnToggleAutoRemove, pressed=" + pressed);
            _autoRemoveEnabled = pressed;
            if (pressed) StartAutoRemove(); else StopAutoRemove();
            _ribbon?.InvalidateControl("AutoRemoveToggle");
        }

        public void OnIntervalChange(IRibbonControl control, string text)
        {
            if (int.TryParse(text, out int ms) && ms >= 100)
            {
                _autoRemoveInterval = ms;
                Logger.Log("Interval changed to " + ms + "ms");
                if (_autoRemoveEnabled)
                {
                    StopAutoRemove();
                    StartAutoRemove();
                }
            }
        }

        public void OnToggleClipboard(IRibbonControl control, bool pressed)
        {
            Logger.Log("OnToggleClipboard, pressed=" + pressed);
            _clipboardEnabled = pressed;
            if (pressed) StartClipboardMonitor(); else StopClipboardMonitor();
            _ribbon?.InvalidateControl("ClipboardToggle");
        }

        public void OnClipboardIntervalChange(IRibbonControl control, string text)
        {
            if (int.TryParse(text, out int ms) && ms >= 100)
            {
                _clipboardInterval = ms;
                Logger.Log("Clipboard interval changed to " + ms + "ms");
                if (_clipboardEnabled)
                {
                    StopClipboardMonitor();
                    StartClipboardMonitor();
                }
            }
        }

        public void OnCopyCleanText(IRibbonControl control)
        {
            Logger.Log("OnCopyCleanText");
            try
            {
                using var oneNote = new OneNoteHelper();
                string text = HyperlinkRemover.GetSelectedText(oneNote);
                if (string.IsNullOrEmpty(text))
                {
                    Logger.Log("No selected text found");
                    return;
                }
                string cleaned = text.Replace(HyperlinkRemover.ZeroWidthSpace, "");
                // Clipboard.SetText requires STA thread
                var thread = new System.Threading.Thread(() =>
                {
                    try { Clipboard.SetText(cleaned); }
                    catch { }
                });
                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                thread.Start();
                thread.Join(500);
                Logger.Log("Copied clean text: " + cleaned.Substring(0, Math.Min(80, cleaned.Length)));
            }
            catch (Exception ex) { Logger.Log("OnCopyCleanText error: " + ex); }
        }

        // Labels
        public string GetTabLabel(IRibbonControl c) => Strings.Get("TabName");
        public string GetGroupLabel(IRibbonControl c) => Strings.Get("GroupName");
        public string GetRemovePageLabel(IRibbonControl c) => Strings.Get("RemovePage");
        public string GetRemoveSelectionLabel(IRibbonControl c) => Strings.Get("RemoveSelection");
        public string GetAutoRemoveLabel(IRibbonControl c) => Strings.Get("AutoRemove");
        public string GetIntervalLabel(IRibbonControl c) => Strings.Get("Interval");
        public string GetClipboardLabel(IRibbonControl c) => Strings.Get("Clipboard");
        public string GetCopyCleanTextLabel(IRibbonControl c) => Strings.Get("CopyCleanText");
        public string GetClipboardIntervalLabel(IRibbonControl c) => Strings.Get("ClipboardInterval");

        // Screentips
        public string GetRemovePageScreentip(IRibbonControl c) => Strings.Get("RemovePageScreen");
        public string GetRemoveSelectionScreentip(IRibbonControl c) => Strings.Get("RemoveSelectionScreen");
        public string GetAutoRemoveScreentip(IRibbonControl c) => Strings.Get("AutoRemoveScreen");
        public string GetIntervalScreentip(IRibbonControl c) => Strings.Get("IntervalScreen");
        public string GetClipboardScreentip(IRibbonControl c) => Strings.Get("ClipboardScreen");
        public string GetCopyCleanTextScreentip(IRibbonControl c) => Strings.Get("CopyCleanTextScreen");
        public string GetClipboardIntervalScreentip(IRibbonControl c) => Strings.Get("ClipboardIntervalScreen");
        public string GetClipboardIntervalText(IRibbonControl control) => _clipboardInterval.ToString();

        // Supertips
        public string GetRemovePageSupertip(IRibbonControl c) => Strings.Get("RemovePageSuper");
        public string GetRemoveSelectionSupertip(IRibbonControl c) => Strings.Get("RemoveSelectionSuper");

        // State getters
        public bool GetAutoRemovePressed(IRibbonControl control) => _autoRemoveEnabled;
        public bool GetClipboardPressed(IRibbonControl control) => _clipboardEnabled;
        public string GetIntervalText(IRibbonControl control) => _autoRemoveInterval.ToString();

        #endregion

        #region Auto-remove

        private void StartAutoRemove()
        {
            if (_autoRemoveTimer != null) return;
            _autoRemoveTimer = new System.Threading.Timer(_ =>
            {
                try
                {
                    using var oneNote = new OneNoteHelper();
                    int count = HyperlinkRemover.RemoveHyperlinksFromCurrentPage(oneNote);
                    if (count > 0) Logger.Log("Auto-remove: " + count);
                }
                catch (Exception ex) { Logger.Log("Auto-remove error: " + ex.Message); }
            }, null, _autoRemoveInterval, _autoRemoveInterval);
            Logger.Log("Auto-remove started, interval=" + _autoRemoveInterval + "ms");
        }

        private void StopAutoRemove()
        {
            _autoRemoveTimer?.Dispose();
            _autoRemoveTimer = null;
        }

        #endregion

        #region Clipboard Monitor

        private void StartClipboardMonitor()
        {
            if (_clipboardMonitor != null) return;
            try { _clipboardMonitor = new ClipboardMonitor(_clipboardInterval); }
            catch (Exception ex) { Logger.Log("ClipboardMonitor error: " + ex.Message); }
        }

        private void StopClipboardMonitor()
        {
            _clipboardMonitor?.Dispose();
            _clipboardMonitor = null;
        }

        #endregion

        #region Helpers

        private void Execute(Func<int> action)
        {
            try
            {
                int count = action();
                Logger.Log("Removed " + count);
            }
            catch (Exception ex) { Logger.Log("Error: " + ex); }
        }

        #endregion
    }
}
