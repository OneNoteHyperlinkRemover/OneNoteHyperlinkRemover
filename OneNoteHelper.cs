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
        private IApplication _application;
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
        /// </summary>
        /// <param name="pageInfo">要获取的页面信息详细程度</param>
        /// <returns>页面 XML 字符串</returns>
        public string GetCurrentPageXml(PageInfo pageInfo = pageInfo.piAll)
        {
            ThrowIfDisposed();

            // 获取当前活动页面的 ID
            string pageId = "";
            _application.GetHierarchy(null, hsSelf, out pageId);

            // 获取页面内容 XML
            string pageXml = "";
            _application.GetPageContent(pageId, out pageXml, pageInfo);
            return pageXml;
        }

        /// <summary>
        /// 获取指定页面的 XML 内容。
        /// </summary>
        public string GetPageXml(string pageId, PageInfo pageInfo = pageInfo.piAll)
        {
            ThrowIfDisposed();

            string pageXml = "";
            _application.GetPageContent(pageId, out pageXml, pageInfo);
            return pageXml;
        }

        /// <summary>
        /// 将修改后的 XML 内容写回页面。
        /// </summary>
        /// <param name="pageId">页面 ID</param>
        /// <param name="pageXml">修改后的页面 XML</param>
        public void UpdatePageContent(string pageId, string pageXml)
        {
            ThrowIfDisposed();

            // lastModified 设为当前时间，期望值设为 0 表示不检查冲突
            _application.UpdatePageContent(pageXml, DateTime.MinValue, xsSchemaCurrent, false);
        }

        /// <summary>
        /// 从页面 XML 中提取页面 ID。
        /// </summary>
        public static string ExtractPageId(string pageXml)
        {
            // 页面 XML 的根元素是 <one:Page>，ID 属性包含页面 ID
            var doc = System.Xml.Linq.XDocument.Parse(pageXml);
            var root = doc.Root;
            if (root == null) return null;

            var ns = root.GetNamespaceOfPrefix(OneNoteNamespacePrefix);
            if (ns == null) return root.Attribute("ID")?.Value;

            return root.Attribute(ns + "ID")?.Value ?? root.Attribute("ID")?.Value;
        }

        #region OneNote COM 常量

        // HierarchyScope
        private const string hsSelf = "Self";

        // XML Schema
        private const XMLSchema xsSchemaCurrent = XMLSchema.xsCurrent;

        #endregion

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
