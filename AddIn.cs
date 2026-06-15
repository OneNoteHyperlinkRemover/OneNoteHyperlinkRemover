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
            try { _clipboardMonitor = new ClipboardMonitor(); }
            catch (Exception ex) { Logger.Log("ClipboardMonitor error: " + ex.Message); }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Logger.Log("=== OnBeginShutdown ===");
            StopAutoRemove();
            _clipboardMonitor?.Dispose();
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

        public bool GetAutoRemovePressed(IRibbonControl control) => _autoRemoveEnabled;

        public string GetGroupLabel(IRibbonControl c) => "超链接工具";
        public string GetButtonLabel(IRibbonControl c) => "移除超链接";
        public string GetAutoRemoveLabel(IRibbonControl c) => "自动移除";

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
            }, null, 1000, 2000);
            Logger.Log("Auto-remove started");
        }

        private void StopAutoRemove()
        {
            _autoRemoveTimer?.Dispose();
            _autoRemoveTimer = null;
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
