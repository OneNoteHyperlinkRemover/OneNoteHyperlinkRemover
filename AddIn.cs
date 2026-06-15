using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// OneNote COM Add-in entry point.
    /// Implements IDTExtensibility2 (COM add-in lifecycle) and IRibbonExtensibility (Ribbon UI).
    /// </summary>
    [ComVisible(true)]
    [Guid("b7a3d2e1-4f6c-4a8b-9e1d-3c5f7a9b2d4e")]
    [ProgId("OneNoteHyperlinkRemover.AddIn")]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private bool _autoRemoveEnabled;
        private System.Threading.Timer _autoRemoveTimer;
        private ClipboardMonitor _clipboardMonitor;
        private static readonly string LogDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OneNoteHyperlinkRemover");
        private static readonly string LogFile = Path.Combine(LogDir, "addin.log");

        public AddIn()
        {
            Log("=== Constructor started ===");
            Log("Assembly location: " + Assembly.GetExecutingAssembly().Location);
            Log("Runtime: " + Environment.Version);
            Log("OS: " + Environment.OSVersion);
        }

        #region IDTExtensibility2

        public void OnConnection(
            object Application,
            ext_ConnectMode ConnectMode,
            object AddInInst,
            ref Array custom)
        {
            Log("=== OnConnection called, ConnectMode=" + ConnectMode + " ===");
            Log("Application type: " + (Application?.GetType().FullName ?? "null"));
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("=== OnDisconnection called, RemoveMode=" + RemoveMode + " ===");
            StopAutoRemove();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            Log("OnAddInsUpdate called");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log("=== OnStartupComplete called ===");
            try
            {
                _clipboardMonitor = new ClipboardMonitor();
                Log("ClipboardMonitor started");
            }
            catch (Exception ex)
            {
                Log("ClipboardMonitor error: " + ex.Message);
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log("=== OnBeginShutdown called ===");
            StopAutoRemove();
            _clipboardMonitor?.Dispose();
            _clipboardMonitor = null;
        }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            Log("=== GetCustomUI called, RibbonID=" + RibbonID + " ===");

            try
            {
                var asm = Assembly.GetExecutingAssembly();
                Log("Embedded resources: " + string.Join(", ", asm.GetManifestResourceNames()));

                using var stream = asm.GetManifestResourceStream("OneNoteHyperlinkRemover.Ribbon.Ribbon.xml");
                if (stream != null)
                {
                    using var reader = new StreamReader(stream);
                    string xml = reader.ReadToEnd();
                    Log("Ribbon XML loaded, length=" + xml.Length);
                    return xml;
                }
                else
                {
                    Log("ERROR: Ribbon.xml resource not found!");
                }
            }
            catch (Exception ex)
            {
                Log("ERROR in GetCustomUI: " + ex);
            }

            return string.Empty;
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Log("=== Ribbon_Load called ===");
        }

        #endregion

        #region Ribbon Callbacks

        public string GetGroupLabel(IRibbonControl control)
        {
            Log("GetGroupLabel called");
            return "超链接工具"; // 超链接工具
        }

        public string GetButtonLabel(IRibbonControl control)
        {
            return "移除超链接"; // 移除超链接
        }

        public string GetButtonScreentip(IRibbonControl control)
        {
            return "移除当前页面的自动超链接"; // 移除当前页面的自动超链接
        }

        public string GetButtonSupertip(IRibbonControl control)
        {
            return "扫描当前 OneNote 页面，将自动转换的 URL 超链接恢复为纯文本。";
        }

        public string GetAutoRemoveLabel(IRibbonControl control)
        {
            return "自动移除"; // 自动移除
        }

        public string GetAutoRemoveScreentip(IRibbonControl control)
        {
            return "开启/关闭自动监控模式"; // 开启/关闭自动监控模式
        }

        public bool GetAutoRemovePressed(IRibbonControl control)
        {
            return _autoRemoveEnabled;
        }

        public stdole.IPictureDisp GetButtonImage(IRibbonControl control)
        {
            Log("GetButtonImage called");
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var stream = asm.GetManifestResourceStream("OneNoteHyperlinkRemover.Resources.RemoveLinks_32.png");
                if (stream != null)
                {
                    var bmp = new System.Drawing.Bitmap(stream);
                    return PictureConverter.Convert(bmp);
                }
            }
            catch (Exception ex)
            {
                Log("GetButtonImage error: " + ex.Message);
            }
            return null;
        }

        public void OnRemoveHyperlinks(IRibbonControl control)
        {
            Log("OnRemoveHyperlinks called");
            try
            {
                using var oneNote = new OneNoteHelper();
                int count = HyperlinkRemover.RemoveHyperlinksFromCurrentPage(oneNote);
                Log("Removed " + count + " hyperlinks");
            }
            catch (Exception ex)
            {
                Log("OnRemoveHyperlinks error: " + ex);
            }
        }

        public void OnToggleAutoRemove(IRibbonControl control, bool pressed)
        {
            Log("OnToggleAutoRemove called, pressed=" + pressed);
            _autoRemoveEnabled = pressed;
            if (pressed)
                StartAutoRemove();
            else
                StopAutoRemove();
            _ribbon?.InvalidateControl("AutoRemoveToggle");
        }

        #endregion

        #region Auto-remove mode

        private void StartAutoRemove()
        {
            if (_autoRemoveTimer != null) return;
            _autoRemoveTimer = new System.Threading.Timer(AutoRemoveTick, null, 1000, 2000);
            Log("Auto-remove started");
        }

        private void StopAutoRemove()
        {
            if (_autoRemoveTimer != null)
            {
                _autoRemoveTimer.Dispose();
                _autoRemoveTimer = null;
                Log("Auto-remove stopped");
            }
        }

        private void AutoRemoveTick(object state)
        {
            try
            {
                Log("Auto-remove tick");
                using var oneNote = new OneNoteHelper();
                int count = HyperlinkRemover.RemoveHyperlinksFromCurrentPage(oneNote);
                if (count > 0)
                    Log("Auto-remove: removed " + count);
            }
            catch (Exception ex)
            {
                Log("Auto-remove tick error: " + ex.Message);
            }
        }

        #endregion

        #region Logging

        private static void Log(string message)
        {
            try
            {
                if (!Directory.Exists(LogDir))
                    Directory.CreateDirectory(LogDir);

                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}";
                File.AppendAllText(LogFile, line);
            }
            catch
            {
                // logging should never break the add-in
            }
        }

        #endregion
    }

    /// <summary>
    /// Converts System.Drawing.Bitmap to stdole.IPictureDisp for Ribbon icons.
    /// </summary>
    internal static class PictureConverter
    {
        public static stdole.IPictureDisp Convert(System.Drawing.Bitmap bitmap)
        {
            return (stdole.IPictureDisp)AxHostHelper.GetIPicture(bitmap);
        }

        private class AxHostHelper : System.Windows.Forms.AxHost
        {
            private AxHostHelper() : base("") { }

            internal static object GetIPicture(System.Drawing.Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }
    }
}
