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
        private bool _clipboardEnabled = true;
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
            Logger.Log($"=== OnDisconnection, mode={mode} ===");
            StopAutoRemove();
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            Logger.Log("=== OnStartupComplete ===");
            if (_clipboardEnabled) StartClipboardMonitor();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Logger.Log("=== OnBeginShutdown ===");
            StopAutoRemove();
            StopClipboardMonitor();
        }

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
                HyperlinkRemover.ClearTracking();
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

        // Labels
        public string GetGroupLabel(IRibbonControl c) => Strings.Get("GroupName");
        public string GetRemovePageLabel(IRibbonControl c) => Strings.Get("RemovePage");
        public string GetRemoveSelectionLabel(IRibbonControl c) => Strings.Get("RemoveSelection");
        public string GetAutoRemoveLabel(IRibbonControl c) => Strings.Get("AutoRemove");
        public string GetIntervalLabel(IRibbonControl c) => Strings.Get("Interval");
        public string GetClipboardLabel(IRibbonControl c) => Strings.Get("Clipboard");

        // Screentips
        public string GetRemovePageScreentip(IRibbonControl c) => Strings.Get("RemovePageScreen");
        public string GetRemoveSelectionScreentip(IRibbonControl c) => Strings.Get("RemoveSelectionScreen");
        public string GetAutoRemoveScreentip(IRibbonControl c) => Strings.Get("AutoRemoveScreen");
        public string GetIntervalScreentip(IRibbonControl c) => Strings.Get("IntervalScreen");
        public string GetClipboardScreentip(IRibbonControl c) => Strings.Get("ClipboardScreen");

        // Supertips
        public string GetRemovePageSupertip(IRibbonControl c) => Strings.Get("RemovePageSuper");
        public string GetRemoveSelectionSupertip(IRibbonControl c) => Strings.Get("RemoveSelectionSuper");

        // State getters
        public bool GetAutoRemovePressed(IRibbonControl control) => _autoRemoveEnabled;
        public bool GetClipboardPressed(IRibbonControl control) => _clipboardEnabled;
        public string GetIntervalText(IRibbonControl control) => _autoRemoveInterval.ToString();

        public stdole.IPictureDisp GetButtonImage(IRibbonControl control)
        {
            try
            {
                using var stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("OneNoteHyperlinkRemover.Resources.RemoveLinks_32.png");
                if (stream != null)
                    return PictureConverter.Convert(new System.Drawing.Bitmap(stream));
            }
            catch { }
            return null;
        }

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
            try { _clipboardMonitor = new ClipboardMonitor(); }
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

    internal static class PictureConverter
    {
        public static stdole.IPictureDisp Convert(System.Drawing.Bitmap bitmap)
            => (stdole.IPictureDisp)AxHostHelper.GetIPicture(bitmap);

        private class AxHostHelper : AxHost
        {
            private AxHostHelper() : base("") { }
            internal static object GetIPicture(System.Drawing.Image image)
                => GetIPictureDispFromPicture(image);
        }
    }
}
