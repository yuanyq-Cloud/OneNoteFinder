using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using Application = Microsoft.Office.Interop.OneNote.Application;

namespace OneFinder
{
    /// <summary>
    /// 代表一个匹配到的 OneNote 页面
    /// </summary>
    public class PageResult
    {
        public string NotebookName { get; set; } = string.Empty;
        public string SectionName  { get; set; } = string.Empty;
        public string PageName     { get; set; } = string.Empty;
        public string PageId       { get; set; } = string.Empty;

        /// <summary>
        /// 搜索词命中的片段列表（每个片段包含前后文）
        /// </summary>
        public List<string> Snippets { get; set; } = new();

        /// <summary>
        /// 命中对象的 ID 列表（用于页内导航）
        /// </summary>
        public List<string> HitObjectIds { get; set; } = new();

        public override string ToString() =>
            $"{NotebookName}  ›  {SectionName}  ›  {PageName}";

        /// <summary>
        /// 获取格式化的搜索结果摘要（包含片段预览）
        /// </summary>
        public string GetDisplayText()
        {
            string basePath = ToString();
            if (Snippets.Count == 0) return basePath;

            // 显示第一个片段
            string firstSnippet = Snippets[0];
            if (Snippets.Count > 1)
                return $"{basePath}\n    {firstSnippet} … (+{Snippets.Count - 1} 处匹配)";
            else
                return $"{basePath}\n    {firstSnippet}";
        }
    }

    /// <summary>
    /// 代表单个匹配项（用于列表显示）
    /// </summary>
    public class MatchResult
    {
        public string NotebookName { get; set; } = string.Empty;
        public string SectionName  { get; set; } = string.Empty;
        public string PageName     { get; set; } = string.Empty;
        public string PageId       { get; set; } = string.Empty;

        /// <summary>
        /// 匹配的文本片段
        /// </summary>
        public string Snippet { get; set; } = string.Empty;

        /// <summary>
        /// 匹配对象的 ID（用于页内导航）
        /// </summary>
        public string? ObjectId { get; set; }

        /// <summary>
        /// 该页面中的匹配序号（1-based）
        /// </summary>
        public int MatchIndex { get; set; }

        /// <summary>
        /// 该页面的总匹配数
        /// </summary>
        public int TotalMatches { get; set; }

        public string GetPagePath() =>
            $"{NotebookName}  ›  {SectionName}  ›  {PageName}";

        public string GetMatchInfo() =>
            TotalMatches > 1 ? $"[{MatchIndex}/{TotalMatches}]" : "";
    }

    /// <summary>
    /// 封装对 OneNote COM API 的访问，通过 XML 全文搜索页面。
    /// </summary>
    public class OneNoteService : IDisposable
    {
        private Application? _app;
        private bool _disposed;

        // OneNote XML 命名空间（版本号为 2013+）
        private static readonly XNamespace NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        public OneNoteService()
        {
            // 如果 OneNote 已经打开则附加，否则启动一个新实例
            _app = new Application();
        }

        /// <summary>
        /// 获取当前打开的笔记本ID
        /// </summary>
        public string? GetCurrentNotebookId()
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            try
            {
                // 获取当前页面ID
                string currentPageId = _app.Windows.CurrentWindow.CurrentPageId;
                if (string.IsNullOrEmpty(currentPageId)) return null;

                // 获取完整层次结构
                _app.GetHierarchy("", HierarchyScope.hsNotebooks, out string xml);
                var doc = XDocument.Parse(xml);

                // 遍历所有笔记本查找包含当前页面的笔记本
                foreach (var notebook in doc.Descendants(NS + "Notebook"))
                {
                    string notebookId = notebook.Attribute("ID")?.Value ?? string.Empty;
                    if (string.IsNullOrEmpty(notebookId)) continue;

                    // 获取笔记本的所有页面
                    _app.GetHierarchy(notebookId, HierarchyScope.hsPages, out string notebookXml);
                    var notebookDoc = XDocument.Parse(notebookXml);

                    // 检查当前页面是否在这个笔记本中
                    var page = notebookDoc.Descendants(NS + "Page")
                        .FirstOrDefault(p => p.Attribute("ID")?.Value == currentPageId);

                    if (page != null)
                    {
                        return notebookId;
                    }
                }
            }
            catch
            {
                // 如果无法获取当前笔记本，返回null
            }

            return null;
        }

        /// <summary>
        /// 遍历所有笔记本 → 节 → 页，在页面 XML 中做大小写不敏感的全文匹配。
        /// </summary>
        /// <param name="query">搜索关键词</param>
        /// <param name="currentNotebookOnly">是否仅搜索当前笔记本</param>
        /// <param name="progress">进度回调</param>
        public List<PageResult> Search(string query, bool currentNotebookOnly = false, Action<string>? progress = null)
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));
            if (string.IsNullOrWhiteSpace(query)) return new List<PageResult>();

            var results = new List<PageResult>();
            string queryLower = query.ToLowerInvariant();

            // 如果需要仅搜索当前笔记本，获取当前笔记本ID
            string? currentNotebookId = null;
            if (currentNotebookOnly)
            {
                currentNotebookId = GetCurrentNotebookId();
                if (string.IsNullOrEmpty(currentNotebookId))
                {
                    progress?.Invoke("无法获取当前笔记本，将搜索所有笔记本");
                    currentNotebookOnly = false;
                }
            }

            // 获取所有笔记本的层次结构 XML
            _app.GetHierarchy(null, HierarchyScope.hsPages, out string hierarchyXml);
            var hierarchy = XDocument.Parse(hierarchyXml);

            foreach (var notebook in hierarchy.Descendants(NS + "Notebook"))
            {
                string nbId = notebook.Attribute("ID")?.Value ?? string.Empty;
                string nbName = notebook.Attribute("name")?.Value ?? "(未命名笔记本)";

                // 如果仅搜索当前笔记本，跳过其他笔记本
                if (currentNotebookOnly && !string.IsNullOrEmpty(currentNotebookId) && nbId != currentNotebookId)
                {
                    continue;
                }

                progress?.Invoke($"扫描笔记本：{nbName}");

                foreach (var section in notebook.Descendants(NS + "Section"))
                {
                    // 跳过受密码保护或已加密的节
                    if (section.Attribute("locked")?.Value == "true") continue;
                    if (section.Attribute("isInRecycleBin")?.Value == "true") continue;

                    string secName = section.Attribute("name")?.Value ?? "(未命名节)";

                    foreach (var page in section.Elements(NS + "Page"))
                    {
                        string pageId   = page.Attribute("ID")?.Value ?? string.Empty;
                        string pageName = page.Attribute("name")?.Value ?? "(未命名页面)";

                        if (string.IsNullOrEmpty(pageId)) continue;

                        try
                        {
                            // 获取页面完整 XML（包含所有文本内容）
                            _app.GetPageContent(pageId, out string pageXml,
                                PageInfo.piAll, XMLSchema.xs2013);

                            var pageDoc = XDocument.Parse(pageXml);
                            var snippets = new List<string>();
                            var hitObjectIds = new List<string>();

                            // 提取所有 T 元素（文本节点）
                            ExtractTextMatches(pageDoc, queryLower, query.Length, snippets, hitObjectIds);

                            if (snippets.Count > 0)
                            {
                                results.Add(new PageResult
                                {
                                    NotebookName = nbName,
                                    SectionName  = secName,
                                    PageName     = pageName,
                                    PageId       = pageId,
                                    Snippets     = snippets,
                                    HitObjectIds = hitObjectIds,
                                });
                            }
                        }
                        catch
                        {
                            // 单页读取失败时跳过，不影响整体搜索
                        }
                    }
                }
            }

            return results;
        }

        /// <summary>
        /// 从页面 XML 中提取包含匹配文本的片段
        /// </summary>
        private void ExtractTextMatches(XDocument pageDoc, string queryLower, int queryLength,
            List<string> snippets, List<string> hitObjectIds)
        {
            // 在 OneNote XML 中，文本存储在 <T> 元素中
            // <OE> 是 Outline Element（大纲元素）
            foreach (var textElement in pageDoc.Descendants(NS + "T"))
            {
                string text = textElement.Value;
                if (string.IsNullOrWhiteSpace(text)) continue;

                string textLower = text.ToLowerInvariant();
                int index = textLower.IndexOf(queryLower);

                if (index >= 0)
                {
                    // 找到匹配，提取上下文片段
                    string snippet = ExtractSnippet(text, index, queryLength);
                    snippets.Add(snippet);

                    // 尝试获取父 OE 元素的 objectID（用于页内导航）
                    var oeElement = textElement.Ancestors(NS + "OE").FirstOrDefault();
                    if (oeElement != null)
                    {
                        string? objectId = oeElement.Attribute("objectID")?.Value;
                        if (!string.IsNullOrEmpty(objectId))
                            hitObjectIds.Add(objectId);
                    }

                    // 如果已经找到足够多的片段，可以提前停止（避免结果过多）
                    if (snippets.Count >= 5) break;
                }
            }
        }

        /// <summary>
        /// 从文本中提取匹配位置的上下文片段
        /// </summary>
        private string ExtractSnippet(string text, int matchIndex, int matchLength, 
            int contextLength = 30)
        {
            // 计算片段的起始和结束位置（在原始文本上）
            int start = Math.Max(0, matchIndex - contextLength);
            int end = Math.Min(text.Length, matchIndex + matchLength + contextLength);

            string prefix = start > 0 ? "…" : "";
            string suffix = end < text.Length ? "…" : "";

            string snippet = text.Substring(start, end - start);

            // 在匹配词前后添加标记（在原始文本上）
            int highlightStart = matchIndex - start;
            int highlightEnd = highlightStart + matchLength;

            if (highlightStart >= 0 && highlightEnd <= snippet.Length)
            {
                snippet = snippet.Substring(0, highlightStart) +
                         "[" + snippet.Substring(highlightStart, matchLength) + "]" +
                         snippet.Substring(highlightEnd);
            }

            // 最后一步：清洗HTML但保留[关键词]标记
            snippet = CleanHtmlContentKeepBrackets(prefix + snippet + suffix);

            return snippet;
        }

        /// <summary>
        /// 清洗HTML标签和实体，但保留[关键词]标记
        /// </summary>
        private string CleanHtmlContentKeepBrackets(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // 临时替换[关键词]标记，避免被清洗破坏
            string placeholder = "\u0001KEYWORD\u0002";  // 使用不可见字符作为占位符
            text = text.Replace("[", placeholder + "[");
            text = text.Replace("]", "]" + placeholder);

            // 用更贪婪和强壮的正则移除HTML闭合标签对比如 <span ...> ... </span> 
            // 有些时候 OneNote 原始 XML 的 T 标签内包装的 CDATA 被当作纯字符串，
            // 里面有不成对或者奇葩结构的碎片，这里做个大清洗。

            // 移除完整的标签
            text = System.Text.RegularExpressions.Regex.Replace(text, @"<([A-Za-z][A-Za-z0-9]*)\b[^>]*>(.*?)</\1>", "$2", System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // 移除所有HTML标签（包括属性）
            // 匹配 <xxx> 或 </xxx> 或 <xxx attr="value"> 等，更强壮地处理跨行和引号内的>
            text = System.Text.RegularExpressions.Regex.Replace(text, @"<[^>]*>", "");

            // 下方这段处理残破的不带 < > 的属性值碎片
            // 某些情况下，OneNote的XML会残留 lang=en-US 等孤立属性，也一并清理
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\b(?:lang|style|class|id|dir|align|title|contenteditable|color|face|size|data-[a-zA-Z0-9\-]+)\s*=\s*""[^""]*""", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\b(?:lang|style|class|id|dir|align|title|contenteditable|color|face|size|data-[a-zA-Z0-9\-]+)\s*=\s*'[^']*'", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\b(?:lang|style|class|id|dir|align|title|contenteditable|color|face|size|data-[a-zA-Z0-9\-]+)\s*=\s*[\w\-]+", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // 移除残留的孤立尖括号
            text = System.Text.RegularExpressions.Regex.Replace(text, @"<|>", "");

            // 替换常见HTML实体
            text = text.Replace("&nbsp;", " ");
            text = text.Replace("&lt;", "<");
            text = text.Replace("&gt;", ">");
            text = text.Replace("&amp;", "&");
            text = text.Replace("&quot;", "\"");
            text = text.Replace("&#39;", "'");
            text = text.Replace("&apos;", "'");

            // 移除其他HTML实体 &#xxx; 或 &#xHHH;
            text = System.Text.RegularExpressions.Regex.Replace(text, @"&#?[a-zA-Z0-9]+;", "");

            // 移除多余空白
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");

            // 恢复[关键词]标记
            text = text.Replace(placeholder, "");

            return text.Trim();
        }

        /// <summary>
        /// 在 OneNote 中打开指定页面（将焦点跳转到该页）。
        /// </summary>
        public void NavigateToPage(string pageId, string? objectId = null)
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            // 如果提供了对象 ID，尝试导航到页内对象
            if (!string.IsNullOrEmpty(objectId))
            {
                try
                {
                    _app.NavigateTo(pageId, objectId);
                    return;
                }
                catch
                {
                    // 如果页内导航失败，降级为普通页面导航
                }
            }

            // 导航到页面
            _app.NavigateTo(pageId);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_app != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
                    _app = null;
                }
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
    }
}
