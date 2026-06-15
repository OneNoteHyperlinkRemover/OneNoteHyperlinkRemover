using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// OneNote COM Add-in 入口点。
    /// 实现 IDTExtensibility2（COM 加载项生命周期）和 IRibbonExtensibility（Ribbon UI）。
    /// 注意：OneNote 不支持 VSTO 模板，因此直接实现 COM Add-in 接口。
    /// 参考项目：OneMore (https://github.com/stevencohn/OneMore)
    /// </summary>
    [ComVisible(true)]
    [Guid("b7a3d2e1-4f6c-4a8b-9e1d-3c5f7a9b2d4e")]
    [ProgId("OneNoteHyperlinkRemover.AddIn")]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private bool _autoRemoveEnabled;
        private Timer _autoRemoveTimer;
        private readonly string _logPath;

        public AddIn()
        {
            _logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OneNoteHyperlinkRemover",
                "addin.log");
        }

        #region IDTExtensibility2 生命周期

        public void OnConnection(
            object Application,
            ext_ConnectMode ConnectMode,
            object AddInInst,
            ref Array custom)
        {
            // 不保存 Application COM 引用，避免阻止 OneNote 正常关闭。
            // 参考 OneMore 的做法：每次操作时创建新的 COM 引用。
            Log("OnConnection: Add-in loaded, ConnectMode=" + ConnectMode);
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection: Add-in unloading, RemoveMode=" + RemoveMode);
            StopAutoRemove();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            // 其他加载项变更时触发，通常无需处理
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete: OneNote startup complete");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown: OneNote shutting down");
            StopAutoRemove();
        }

        #endregion

        #region IRibbonExtensibility Ribbon UI

        public string GetCustomUI(string RibbonID)
        {
            // RibbonID: "OneNote" for the main ribbon
            Log("GetCustomUI: RibbonID=" + RibbonID);

            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using var stream = asm.GetManifestResourceStream("OneNoteHyperlinkRemover.Ribbon.Ribbon.xml");
                if (stream != null)
                {
                    using var reader = new StreamReader(stream);
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Log("GetCustomUI error: " + ex.Message);
            }

            return string.Empty;
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Log("Ribbon loaded");
        }

        #endregion

        #region Ribbon 回调方法

        public string GetGroupLabel(IRibbonControl control)
        {
            return "超链接工具";
        }

        public string GetButtonLabel(IRibbonControl control)
        {
            return "移除超链接";
        }

        public string GetButtonScreentip(IRibbonControl control)
        {
            return "移除当前页面的自动超链接";
        }

        public string GetButtonSupertip(IRibbonControl control)
        {
            return "扫描当前 OneNote 页面，将自动转换的 URL 超链接恢复为纯文本。不会影响手动插入的超链接。";
        }

        public string GetAutoRemoveLabel(IRibbonControl control)
        {
            return "自动移除";
        }

        public string GetAutoRemoveScreentip(IRibbonControl control)
        {
            return "开启/关闭自动监控模式";
        }

        public bool GetAutoRemovePressed(IRibbonControl control)
        {
            return _autoRemoveEnabled;
        }

        public stdole.IPictureDisp GetButtonImage(IRibbonControl control)
        {
            // 尝试从嵌入资源加载图标
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
            catch
            {
                // 图标加载失败时返回 null，Ribbon 会显示默认图标
            }
            return null;
        }

        public void OnRemoveHyperlinks(IRibbonControl control)
        {
            try
            {
                using var oneNote = new OneNoteHelper();
                int count = HyperlinkRemover.RemoveHyperlinksFromCurrentPage(oneNote);
                if (count > 0)
                {
                    MessageBox.Show(
                        $"已移除 {count} 个自动超链接。",
                        "移除完成",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "当前页面没有找到需要移除的自动超链接。",
                        "无需处理",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Log("OnRemoveHyperlinks error: " + ex);
                MessageBox.Show(
                    "移除超链接时出错：\n" + ex.Message,
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void OnToggleAutoRemove(IRibbonControl control, bool pressed)
        {
            _autoRemoveEnabled = pressed;
            if (pressed)
            {
                StartAutoRemove();
            }
            else
            {
                StopAutoRemove();
            }
            _ribbon?.InvalidateControl("AutoRemoveToggle");
        }

        #endregion

        #region 自动移除模式

        private void StartAutoRemove()
        {
            if (_autoRemoveTimer != null) return;

            _autoRemoveTimer = new Timer { Interval = 2000 }; // 每 2 秒检查一次
            _autoRemoveTimer.Tick += AutoRemoveTick;
            _autoRemoveTimer.Start();
            Log("Auto-remove started");
        }

        private void StopAutoRemove()
        {
            if (_autoRemoveTimer != null)
            {
                _autoRemoveTimer.Stop();
                _autoRemoveTimer.Dispose();
                _autoRemoveTimer = null;
                Log("Auto-remove stopped");
            }
        }

        private void AutoRemoveTick(object sender, EventArgs e)
        {
            try
            {
                using var oneNote = new OneNoteHelper();
                HyperlinkRemover.RemoveHyperlinksFromCurrentPage(oneNote);
            }
            catch (Exception ex)
            {
                Log("Auto-remove tick error: " + ex.Message);
            }
        }

        #endregion

        #region 日志

        private void Log(string message)
        {
            try
            {
                var dir = Path.GetDirectoryName(_logPath);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                File.AppendAllText(_logPath, line);
            }
            catch
            {
                // 日志写入失败不应影响主功能
            }
        }

        #endregion
    }

    /// <summary>
    /// 将 System.Drawing.Bitmap 转换为 stdole.IPictureDisp（Ribbon 图标所需格式）。
    /// </summary>
    internal static class PictureConverter
    {
        public static stdole.IPictureDisp Convert(System.Drawing.Bitmap bitmap)
        {
            return (stdole.IPictureDisp)AxHostConverter.GetIPictureDispFromPicture(bitmap);
        }

        private class AxHostConverter : System.Windows.Forms.AxHost
        {
            private AxHostConverter() : base("") { }

            public static object GetIPictureDispFromPicture(System.Drawing.Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }
    }
}
