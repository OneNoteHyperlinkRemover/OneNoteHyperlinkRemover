using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteHyperlinkRemover
{
    /// <summary>
    /// OneNote COM API 封装类。
    /// 管理 IApplication COM 对象的生命周期，提供页面读写操作。
    /// 关键设计：不长期持有 COM 引用，每次操作后及时释放，避免阻止 OneNote 关闭。
    /// </summary>
    internal sealed class OneNoteHelper : IDisposable
    {
        private Application _application;
        private bool _disposed;

        // OneNote 页面 XML 的命名空间前缀
        public const string OneNoteNamespacePrefix = "one";
        public const string OneNoteNamespace = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        public OneNoteHelper()
        {
            _application = new Application();
        }

        /// <summary>
        /// 获取当前活动页面的完整 XML 内容。
        /// 通过 Windows.CurrentWindow.CurrentPageId 获取当前页面 ID。
        /// </summary>
        public string GetCurrentPageXml(PageInfo pageInfo = PageInfo.piAll)
        {
            ThrowIfDisposed();

            // 通过 Windows 集合获取当前活动窗口的页面 ID（参考 OneMore）
            string pageId = GetCurrentPageId();
            if (string.IsNullOrEmpty(pageId))
                throw new InvalidOperationException("无法获取当前页面 ID，请确保 OneNote 中有打开的页面。");

            string pageXml;
            _application.GetPageContent(pageId, out pageXml, pageInfo);
            return pageXml;
        }

        /// <summary>
        /// 获取当前活动页面的 ID。
        /// </summary>
        private string GetCurrentPageId()
        {
            Windows windows = null;
            Window window = null;
            try
            {
                windows = _application.Windows;
                window = windows.CurrentWindow;
                return window?.CurrentPageId;
            }
            finally
            {
                if (window != null && Marshal.IsComObject(window))
                    Marshal.ReleaseComObject(window);
                if (windows != null && Marshal.IsComObject(windows))
                    Marshal.ReleaseComObject(windows);
            }
        }

        /// <summary>
        /// 获取指定页面的 XML 内容。
        /// </summary>
        public string GetPageXml(string pageId, PageInfo pageInfo = PageInfo.piAll)
        {
            ThrowIfDisposed();

            string pageXml;
            _application.GetPageContent(pageId, out pageXml, pageInfo);
            return pageXml;
        }

        /// <summary>
        /// 将修改后的 XML 内容写回页面。
        /// </summary>
        public void UpdatePageContent(string pageId, string pageXml)
        {
            ThrowIfDisposed();

            _application.UpdatePageContent(pageXml, DateTime.MinValue, XMLSchema.xs2013, true);
        }

        /// <summary>
        /// 从页面 XML 中提取页面 ID。
        /// </summary>
        public static string ExtractPageId(string pageXml)
        {
            var doc = System.Xml.Linq.XDocument.Parse(pageXml);
            var root = doc.Root;
            if (root == null) return null;

            var ns = root.GetNamespaceOfPrefix(OneNoteNamespacePrefix);
            if (ns == null) return root.Attribute("ID")?.Value;

            return root.Attribute(ns + "ID")?.Value ?? root.Attribute("ID")?.Value;
        }

        #region IDisposable

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                if (_application != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(_application);
                    }
                    catch
                    {
                        // 忽略释放错误
                    }
                    _application = null;
                }
            }
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(OneNoteHelper));
        }

        #endregion
    }
}
