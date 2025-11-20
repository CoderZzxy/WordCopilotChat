using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using WordCopilotChat.utils;


// 使用别名解决命名冲突
using TaskAsync = System.Threading.Tasks.Task;

namespace WordCopilot.utils
{
    /// <summary>
    /// Word文档操作工具类，为AI Agent提供Word相关功能
    /// </summary>
    public static class WordTools
    {
        private static MarkdownToWord _markdownToWord = new MarkdownToWord();

        // 防止重复执行样式修改的锁
        private static readonly object _styleModificationLock = new object();
        private static bool _isStyleModificationInProgress = false;

        // 用于取消长时间运行操作的标志
        private static volatile bool _shouldCancelOperation = false;

        // 控制是否启用范围清理功能（可以通过此开关禁用以避免异常）
        private static bool _enableRangeNormalization = true;

        // 默认插入位置设置
        private static string _defaultInsertPosition = "end";

        // 是否显示工具调用信息
        private static bool _showToolCalls = true;

        // 是否显示调试详细信息
        private static bool _showDebugInfo = false;

        // 工具进度通知辅助方法
        private static void NotifyProgress(string message)
        {
            Debug.WriteLine(message);

            // 判断消息类型
            bool isToolCall = message.Contains("轮工具调用") || message.Contains("执行工具:") || message.Contains("执行完成");
            bool isDebugInfo = message.Contains("查找标题") || message.Contains("提取关键词") || message.Contains("选择最佳匹配") ||
                              message.Contains("找到候选标题") || message.Contains("标题搜索完成") || message.Contains("提取内容完成") ||
                              message.Contains("连接到已打开的Word实例") || message.Contains("开始高效查找标题") ||
                              message.Contains("开始提取标题下的内容") || message.Contains("找到标题:");

            // 根据配置决定是否发送
            bool shouldSend = (isToolCall && _showToolCalls) || (isDebugInfo && _showDebugInfo);

            if (shouldSend)
            {
                // 通过OpenAIUtils发送进度通知
                OpenAIUtils.NotifyToolProgress(message);
            }
        }

        // 更新Agent配置
        public static void UpdateAgentConfig(bool enableRangeNormalization, string defaultInsertPosition, bool showToolCalls, bool showDebugInfo)
        {
            _enableRangeNormalization = enableRangeNormalization;
            _defaultInsertPosition = NormalizeInsertPosition(defaultInsertPosition);
            _showToolCalls = showToolCalls;
            _showDebugInfo = showDebugInfo;

            Debug.WriteLine($"WordTools配置已更新:");
            Debug.WriteLine($"- 启用范围清理: {_enableRangeNormalization}");
            Debug.WriteLine($"- 默认插入位置: {_defaultInsertPosition}");
            Debug.WriteLine($"- 显示工具调用信息: {_showToolCalls}");
            Debug.WriteLine($"- 显示调试详细信息: {_showDebugInfo}");
        }

        // 规范化插入位置，任何非法值都回退为 end
        private static string NormalizeInsertPosition(string pos)
        {
            try
            {
                string p = (pos ?? "end").Trim().ToLowerInvariant();
                return p == "front" ? "front" : "end";
            }
            catch
            {
                return "end";
            }
        }

        // 辅助方法：安全获取字典值
        private static T GetParameterValue<T>(Dictionary<string, object> parameters, string key, T defaultValue = default(T))
        {
            if (parameters.ContainsKey(key) && parameters[key] != null)
            {
                try
                {
                    if (typeof(T) == typeof(string))
                    {
                        return (T)(object)(parameters[key].ToString() ?? defaultValue.ToString());
                    }
                    return (T)Convert.ChangeType(parameters[key], typeof(T));
                }
                catch
                {
                    return defaultValue;
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// 设置取消操作标志
        /// </summary>
        public static void CancelCurrentOperation()
        {
            _shouldCancelOperation = true;
            Debug.WriteLine("WordTools: 设置取消操作标志");
        }

        /// <summary>
        /// 重置取消操作标志
        /// </summary>
        public static void ResetCancelFlag()
        {
            _shouldCancelOperation = false;
        }

        /// <summary>
        /// 检查是否应该取消当前操作
        /// </summary>
        private static bool ShouldCancelOperation()
        {
            return _shouldCancelOperation;
        }

        /// <summary>
        /// 注册所有Word工具到OpenAIUtils
        /// </summary>
        public static void RegisterAllTools()
        {
            // 清空现有工具
            OpenAIUtils.ClearTools();

            // 注册获取选中文本工具
            RegisterGetSelectedTextTool();

            // 注册格式化插入工具
            RegisterFormattedInsertTool();

            // 注册文字样式修改工具
            RegisterTextStyleTool();

            // 注册文档统计工具
            RegisterDocumentStatsTool();

            // 注册获取图片工具
            RegisterGetImagesTools();

            // 注册获取公式工具
            RegisterGetFormulasTools();

            // 注册获取表格工具
            RegisterGetTablesTools();

            // 注册获取文档标题列表工具（整体结构）
            RegisterGetDocumentHeadingsTool();

            // 注册获取标题内容工具
            RegisterGetHeadingContentTool();

            // 注册检查插入位置工具
            RegisterCheckInsertPositionTool();

            Debug.WriteLine($"Word工具注册完成，共注册 {OpenAIUtils.GetToolCount()} 个工具");
        }

        /// <summary>
        /// 根据启用的工具列表注册选定的工具
        /// </summary>
        /// <param name="enabledTools">要启用的工具名称数组</param>
        public static void RegisterSelectedTools(string[] enabledTools)
        {
            // 清空现有工具
            OpenAIUtils.ClearTools();

            foreach (string toolName in enabledTools)
            {
                switch (toolName)
                {
                    case "get_selected_text":
                        RegisterGetSelectedTextTool();
                        break;
                    case "formatted_insert_content":
                        RegisterFormattedInsertTool();
                        break;
                    case "modify_text_style":
                        RegisterTextStyleTool();
                        break;
                    case "get_document_statistics":
                        RegisterDocumentStatsTool();
                        break;
                    case "get_document_images":
                        RegisterGetImagesTools();
                        break;
                    case "get_document_formulas":
                        RegisterGetFormulasTools();
                        break;
                    case "get_document_tables":
                        RegisterGetTablesTools();
                        break;
                    case "get_document_headings":
                        RegisterGetDocumentHeadingsTool();
                        break;
                    case "get_heading_content":
                        RegisterGetHeadingContentTool();
                        break;
                    case "check_insert_position":
                        RegisterCheckInsertPositionTool();
                        break;
                    default:
                        Debug.WriteLine($"未知工具: {toolName}");
                        break;
                }
            }

            Debug.WriteLine($"根据用户选择注册了 {OpenAIUtils.GetToolCount()} 个工具: {string.Join(", ", enabledTools)}");
        }

        /// <summary>
        /// 获取文档标题，专门用于快捷选择器（支持分页和取消）
        /// 从Word导航窗格数据中获取，确保与导航窗格显示一致
        /// </summary>
        /// <param name="page">页码（从0开始）</param>
        /// <param name="pageSize">每页数量（默认10条）</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>JSON格式的分页标题数据</returns>
        public static string GetDocumentHeadingsForQuickSelector(int page = 0, int pageSize = 10, CancellationToken cancellationToken = default)
        {
            try
            {
                Debug.WriteLine($"开始获取文档标题 - 页码: {page}, 每页: {pageSize}");

                // 获取Word应用程序实例
                dynamic wordApp = _markdownToWord.GetWordApplication();
                if (wordApp == null)
                {
                    Debug.WriteLine("无法连接到Word应用程序");
                    return JsonConvert.SerializeObject(new { headings = new object[0], hasMore = false, total = 0 });
                }

                dynamic doc = _markdownToWord.GetActiveDocument(wordApp);
                if (doc == null)
                {
                    Debug.WriteLine("无法获取活动文档");
                    return JsonConvert.SerializeObject(new { headings = new object[0], hasMore = false, total = 0 });
                }

                // 如果pageSize很大（>1000），说明用户希望一次性加载所有标题，使用简化的快速方法
                if (pageSize > 1000)
                {
                    Debug.WriteLine("检测到大pageSize，使用快速一次性获取所有标题...");
                    var startTime = System.Diagnostics.Stopwatch.StartNew();
                    
                    var allHeadings = GetHeadingsUsingOutlineLevel(doc, cancellationToken);
                    
                    startTime.Stop();
                    Debug.WriteLine($"快速获取完成 - 总标题数: {allHeadings.Count}, 耗时: {startTime.ElapsedMilliseconds}ms");

                    var result = new
                    {
                        headings = allHeadings,
                        hasMore = false,
                        total = allHeadings.Count,
                        page = 0,
                        pageSize = allHeadings.Count
                    };

                    return JsonConvert.SerializeObject(result, Formatting.None);
                }

                // 正常分页获取
                Debug.WriteLine("使用OutlineLevel属性分页获取标题...");
                var pagedResult = GetHeadingsUsingOutlineLevelPaged(doc, page, pageSize, cancellationToken);

                // 检查是否被取消
                cancellationToken.ThrowIfCancellationRequested();

                // 显式解构ValueTuple
                List<object> pagedHeadings = pagedResult.Item1;
                int totalCount = pagedResult.Item2;
                bool hasMore = pagedResult.Item3;

                Debug.WriteLine($"分页结果 - 总数: {totalCount}, 当前页: {pagedHeadings.Count}, 还有更多: {hasMore}");

                var result2 = new
                {
                    headings = pagedHeadings,
                    hasMore = hasMore,
                    total = totalCount,
                    page = page,
                    pageSize = pageSize
                };

                return JsonConvert.SerializeObject(result2, Formatting.None);
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("获取文档标题被取消");
                throw; // 重新抛出取消异常
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档标题时出错: {ex.Message}");
                return JsonConvert.SerializeObject(new { headings = new object[0], hasMore = false, total = 0, error = ex.Message });
            }
        }



        /// <summary>
        /// 使用OutlineLevel属性分页获取标题（真正的分页，只获取需要的数据）
        /// </summary>
        private static (List<object> headings, int totalCount, bool hasMore) GetHeadingsUsingOutlineLevelPaged(dynamic doc, int page, int pageSize, CancellationToken cancellationToken = default)
        {
            var pagedHeadings = new List<object>();
            int startIndex = page * pageSize;
            int endIndex = startIndex + pageSize;

            try
            {
                Debug.WriteLine($"⏱️ 开始高效分页扫描标题 - 页码: {page}, 每页: {pageSize}, 需要索引: {startIndex}-{endIndex}");
                var scanStartTime = System.Diagnostics.Stopwatch.StartNew();

                // 使用缓存的标题总数（如果可用）
                int totalCount = GetCachedHeadingCount(doc);
                if (totalCount == -1)
                {
                    // 首次扫描，快速计算总数
                    totalCount = CountHeadingsQuickly(doc, cancellationToken);
                    CacheHeadingCount(doc, totalCount);
                }

                Debug.WriteLine($"⏱️ 标题总数获取完成: {totalCount}, 耗时: {scanStartTime.ElapsedMilliseconds}ms");

                // 如果请求的页码超出范围，返回空结果
                if (startIndex >= totalCount)
                {
                    Debug.WriteLine($"⏱️ 请求页码超出范围，返回空结果");
                    return (pagedHeadings, totalCount, false);
                }

                // 高效获取指定范围的标题
                int currentIndex = 0;
                int foundCount = 0;
                int processedCount = 0;

                foreach (dynamic para in doc.Paragraphs)
                {
                    try
                    {
                        // 每处理100个段落检查一次取消请求和性能
                        if (processedCount % 100 == 0)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (processedCount > 0 && scanStartTime.ElapsedMilliseconds > 10000) // 超过10秒强制停止
                            {
                                Debug.WriteLine($"⏱️ 扫描超时，强制停止。已处理: {processedCount}个段落");
                                break;
                            }
                        }
                        processedCount++;

                        // 安全检查OutlineLevel
                        try
                        {
                            var outlineLevel = para.OutlineLevel;
                            if (outlineLevel != null && outlineLevel >= 1 && outlineLevel <= 9)
                            {
                                // 只有在目标范围内才获取详细信息
                                if (currentIndex >= startIndex && currentIndex < endIndex)
                                {
                                    try
                                    {
                                        string text = para.Range?.Text?.Trim() ?? "";
                                        if (!string.IsNullOrEmpty(text))
                                        {
                                            text = text.Replace("\r", "").Replace("\n", "").Trim();

                                            if (!string.IsNullOrEmpty(text))
                                            {
                                                // 延迟获取页码（最耗时的操作）
                                                var pageNum = GetPageNumberSafe(para.Range);
                                                var headingInfo = new
                                                {
                                                    text = text,
                                                    level = (int)outlineLevel,
                                                    page = pageNum
                                                };

                                                pagedHeadings.Add(headingInfo);
                                                foundCount++;

                                                // 如果已经找到足够的标题，可以提前退出
                                                if (foundCount >= pageSize)
                                                {
                                                    Debug.WriteLine($"⏱️ 已找到足够标题({foundCount}个)，提前退出扫描");
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception textEx)
                                    {
                                        Debug.WriteLine($"⏱️ 获取标题文本时出错: {textEx.Message}");
                                        continue;
                                    }
                                }

                                currentIndex++;

                                // 如果已经超过了需要的范围，且不是第一页，可以提前退出
                                if (currentIndex > endIndex && page > 0)
                                {
                                    Debug.WriteLine($"⏱️ 超出目标范围，提前退出扫描");
                                    break;
                                }
                            }
                        }
                        catch (Exception levelEx)
                        {
                            Debug.WriteLine($"⏱️ 检查OutlineLevel时出错: {levelEx.Message}");
                            continue;
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        Debug.WriteLine("⏱️ 分页标题扫描被取消");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"⏱️ 处理段落时出错: {ex.Message}");
                        continue;
                    }
                }

                bool hasMore = endIndex < totalCount;
                scanStartTime.Stop();

                Debug.WriteLine($"⏱️ 高效分页扫描完成 - 总标题数: {totalCount}, 当前页标题数: {pagedHeadings.Count}, 还有更多: {hasMore}, 总耗时: {scanStartTime.ElapsedMilliseconds}ms");
                return (pagedHeadings, totalCount, hasMore);
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("⏱️ 分页OutlineLevel扫描被取消");
                throw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⏱️ 分页OutlineLevel方法出错: {ex.Message}");
                // 返回空结果而不是崩溃
                return (pagedHeadings, 0, false);
            }
        }

        // 标题计数缓存（避免重复扫描）
        private static Dictionary<string, int> _headingCountCache = new Dictionary<string, int>();

        /// <summary>
        /// 获取缓存的标题总数
        /// </summary>
        private static int GetCachedHeadingCount(dynamic doc)
        {
            try
            {
                string docKey = GetDocumentKey(doc);
                return _headingCountCache.ContainsKey(docKey) ? _headingCountCache[docKey] : -1;
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// 缓存标题总数
        /// </summary>
        private static void CacheHeadingCount(dynamic doc, int count)
        {
            try
            {
                string docKey = GetDocumentKey(doc);
                _headingCountCache[docKey] = count;

                // 限制缓存大小，避免内存泄漏
                if (_headingCountCache.Count > 10)
                {
                    var firstKey = _headingCountCache.Keys.First();
                    _headingCountCache.Remove(firstKey);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⏱️ 缓存标题总数失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取文档唯一标识
        /// </summary>
        private static string GetDocumentKey(dynamic doc)
        {
            try
            {
                // 使用文档名称和修改时间作为唯一标识
                string name = doc.Name ?? "Unknown";
                string path = doc.FullName ?? "";
                return $"{name}_{path}_{DateTime.Now.Ticks / 10000000}"; // 精确到秒
            }
            catch
            {
                return $"doc_{DateTime.Now.Ticks}";
            }
        }

        /// <summary>
        /// 快速计算标题总数（不获取详细信息）
        /// </summary>
        private static int CountHeadingsQuickly(dynamic doc, CancellationToken cancellationToken = default)
        {
            int count = 0;
            int processedCount = 0;
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                Debug.WriteLine($"⏱️ 开始快速计算标题总数...");

                foreach (dynamic para in doc.Paragraphs)
                {
                    try
                    {
                        // 每处理200个段落检查一次
                        if (processedCount % 200 == 0)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            // 超过5秒强制停止计数
                            if (stopwatch.ElapsedMilliseconds > 5000)
                            {
                                Debug.WriteLine($"⏱️ 计数超时，强制停止。已处理: {processedCount}个段落，找到: {count}个标题");
                                break;
                            }
                        }
                        processedCount++;

                        // 安全检查OutlineLevel，不获取文本内容
                        try
                        {
                            var outlineLevel = para.OutlineLevel;
                            if (outlineLevel != null && outlineLevel >= 1 && outlineLevel <= 9)
                            {
                                // 快速检查是否有文本内容（不获取具体内容）
                                try
                                {
                                    if (para.Range?.Text?.Length > 1) // 至少有一个字符加换行符
                                    {
                                        count++;
                                    }
                                }
                                catch (Exception textEx)
                                {
                                    // 如果无法获取文本，假设有内容（保守估计）
                                    count++;
                                    Debug.WriteLine($"⏱️ 计数时获取文本长度出错，假设有内容: {textEx.Message}");
                                }
                            }
                        }
                        catch (Exception levelEx)
                        {
                            Debug.WriteLine($"⏱️ 计数时检查OutlineLevel出错: {levelEx.Message}");
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"⏱️ 计数时处理段落出错: {ex.Message}");
                        continue;
                    }
                }

                stopwatch.Stop();
                Debug.WriteLine($"⏱️ 快速计数完成 - 总标题数: {count}, 处理段落: {processedCount}, 耗时: {stopwatch.ElapsedMilliseconds}ms");
                return count;
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine($"⏱️ 快速计数被取消 - 已找到: {count}个标题");
                throw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⏱️ 快速计数出错: {ex.Message}");
                return count; // 返回已计算的数量
            }
        }

        /// <summary>
        /// 安全获取页码（优化版本，增强异常处理）
        /// </summary>
        private static int GetPageNumberSafe(dynamic range)
        {
            try
            {
                if (range == null) return 1;

                // 方法1：使用数字索引（最快）
                try
                {
                    var pageInfo = range.Information[7]; // wdActiveEndPageNumber = 7
                    if (pageInfo != null && pageInfo is int)
                    {
                        return (int)pageInfo;
                    }
                }
                catch (Exception ex1)
                {
                    Debug.WriteLine($"⏱️ 页码获取方法1失败: {ex1.Message}");
                }

                // 方法2：使用枚举常量
                try
                {
                    var pageInfo = range.Information[Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber];
                    if (pageInfo != null && pageInfo is int)
                    {
                        return (int)pageInfo;
                    }
                }
                catch (Exception ex2)
                {
                    Debug.WriteLine($"⏱️ 页码获取方法2失败: {ex2.Message}");
                }

                // 方法3：使用Start属性估算页码（备用方法）
                try
                {
                    var start = range.Start;
                    if (start != null && start is int)
                    {
                        // 粗略估算：每页约500个字符
                        return Math.Max(1, ((int)start / 500) + 1);
                    }
                }
                catch (Exception ex3)
                {
                    Debug.WriteLine($"⏱️ 页码获取方法3失败: {ex3.Message}");
                }

                return 1; // 所有方法都失败时返回默认页码
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⏱️ 获取页码时发生未预期异常: {ex.Message}");
                return 1;
            }
        }

        /// <summary>
        /// 快速获取页码（保留原方法名以兼容）
        /// </summary>
        private static int GetPageNumberFast(dynamic range)
        {
            return GetPageNumberSafe(range);
        }

        /// <summary>
        /// 使用OutlineLevel属性获取标题（高效且准确，支持取消）
        /// </summary>
        private static List<object> GetHeadingsUsingOutlineLevel(dynamic doc, CancellationToken cancellationToken = default)
        {
            var headings = new List<object>();

            try
            {
                Debug.WriteLine("开始使用OutlineLevel属性扫描标题...");

                int processedCount = 0;
                foreach (dynamic para in doc.Paragraphs)
                {
                    try
                    {
                        // 每处理50个段落检查一次取消请求
                        if (processedCount % 50 == 0)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                        }
                        processedCount++;

                        // 直接检查OutlineLevel属性
                        var outlineLevel = para.OutlineLevel;
                        if (outlineLevel != null && outlineLevel >= 1 && outlineLevel <= 9)
                        {
                            string text = para.Range.Text?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(text))
                            {
                                text = text.Replace("\r", "").Replace("\n", "").Trim();

                                if (!string.IsNullOrEmpty(text))
                                {
                                    var pageNum = GetPageNumber(para.Range);
                                    var headingInfo = new
                                    {
                                        text = text,
                                        level = (int)outlineLevel,
                                        page = pageNum
                                    };

                                    headings.Add(headingInfo);
                                    Debug.WriteLine($"OutlineLevel标题: {text} (级别{outlineLevel}, 页码{pageNum})");
                                }
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        Debug.WriteLine("标题扫描被取消");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"处理段落OutlineLevel时出错: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("OutlineLevel扫描被取消");
                throw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OutlineLevel方法出错: {ex.Message}");
            }

            Debug.WriteLine($"OutlineLevel方法找到 {headings.Count} 个标题");
            return headings;
        }



        /// <summary>
        /// 从样式名称获取标题级别
        /// </summary>
        /// <param name="styleName">样式名称</param>
        /// <returns>标题级别（1-9），非标题返回0</returns>
        private static int GetHeadingLevel(string styleName)
        {
            if (string.IsNullOrEmpty(styleName))
                return 0;

            // 标准化样式名称（去除空格，转为小写进行比较）
            string normalizedStyle = styleName.Trim().ToLower();

            // 中文标题样式的各种变体
            string[] chineseHeadingPatterns = {
                "标题", "标题 ", "title", "heading"
            };

            foreach (string pattern in chineseHeadingPatterns)
            {
                if (normalizedStyle.StartsWith(pattern.ToLower()))
                {
                    // 提取级别数字
                    string levelPart = styleName.Substring(pattern.Length).Trim();

                    // 处理各种可能的级别表示方式
                    levelPart = levelPart.Replace(" ", "").Replace("级", "").Replace("level", "");

                    if (int.TryParse(levelPart, out int level) && level >= 1 && level <= 9)
                    {
                        return level;
                    }
                }
            }

            // 英文标题样式的各种变体
            string[] englishHeadingPatterns = {
                "heading", "head", "h"
            };

            foreach (string pattern in englishHeadingPatterns)
            {
                if (normalizedStyle.StartsWith(pattern))
                {
                    string levelPart = styleName.Substring(pattern.Length).Trim();
                    levelPart = levelPart.Replace(" ", "").Replace("level", "");

                    if (int.TryParse(levelPart, out int level) && level >= 1 && level <= 9)
                    {
                        return level;
                    }
                }
            }

            // 检查是否包含标题相关的关键词和数字
            if (normalizedStyle.Contains("标题") || normalizedStyle.Contains("heading") ||
                normalizedStyle.Contains("title") || normalizedStyle.Contains("head"))
            {
                // 尝试从样式名称中提取数字
                for (int i = 1; i <= 9; i++)
                {
                    if (normalizedStyle.Contains(i.ToString()))
                    {
                        return i;
                    }
                }
            }

            // 特殊情况：直接检查是否是纯数字标题样式（如"1", "2", "3"等）
            if (styleName.Length == 1 && char.IsDigit(styleName[0]))
            {
                if (int.TryParse(styleName, out int directLevel) && directLevel >= 1 && directLevel <= 9)
                {
                    return directLevel;
                }
            }

            return 0;
        }

        // 注册格式化插入工具
        private static void RegisterFormattedInsertTool()
        {
            OpenAIUtils.RegisterTool(
                "formatted_insert_content",
                "插入格式化的内容到Word文档。可以在指定标题下方插入，也可以在当前光标位置插入",
                new Dictionary<string, object>
                {
                    { "target_heading", new { type = "string", description = "目标标题文本（可选）。如果指定，将在此标题下方插入内容；如果为空，将在当前光标位置插入", @default = "" } },
                    { "content", new { type = "string", description = "要插入的内容" } },
                    { "format_type", new { type = "string", description = "内容格式类型：paragraph(段落), list(列表), table(表格), emphasis(强调)", @default = "paragraph" } },
                    { "indent_level", new { type = "integer", description = "缩进级别(0-5)，用于层次化显示", @default = 0 } },
                    { "add_spacing", new { type = "boolean", description = "是否在插入内容前后添加间距", @default = true } },
                    { "insert_position", new { type = "string", description = "插入位置：front(标题下最前面), end(标题下已有内容的末尾，默认)", @default = "end" } }
                },
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => FormattedInsertContent(parameters));
                });
        }

        // 注册文字样式修改工具
        private static void RegisterTextStyleTool()
        {
            OpenAIUtils.RegisterTool(
                "modify_text_style",
                "修改指定文本的样式，包括背景色、字体大小、粗细、段落间距等格式",
                new Dictionary<string, object>
                {
                    { "text_to_find", new { type = "string", description = "要查找和修改的文本内容" } },
                    { "font_size", new { type = "integer", description = "字体大小(8-72)，单位为磅", minimum = 8, maximum = 72 } },
                    { "font_bold", new { type = "boolean", description = "是否加粗" } },
                    { "font_italic", new { type = "boolean", description = "是否斜体" } },
                    { "font_color", new { type = "string", description = "字体颜色：red, blue, green, black, white, yellow, orange, purple, gray", @enum = new[] { "red", "blue", "green", "black", "white", "yellow", "orange", "purple", "gray" } } },
                    { "background_color", new { type = "string", description = "背景颜色：yellow, lightblue, lightgreen, pink, lightgray, white, none", @enum = new[] { "yellow", "lightblue", "lightgreen", "pink", "lightgray", "white", "none" } } },
                    { "paragraph_spacing_before", new { type = "integer", description = "段落前间距(0-100)，单位为磅", minimum = 0, maximum = 100 } },
                    { "paragraph_spacing_after", new { type = "integer", description = "段落后间距(0-100)，单位为磅", minimum = 0, maximum = 100 } },
                    { "line_spacing", new { type = "number", description = "行间距倍数(1.0-3.0)", minimum = 1.0, maximum = 3.0 } }
                },
                async (parameters) =>
                {
                    // 直接在当前线程执行，避免线程切换导致的COM问题
                    return await TaskAsync.FromResult(ModifyTextStyle(parameters));
                });
        }



        // 格式化插入内容实现
        private static string FormattedInsertContent(Dictionary<string, object> parameters)
        {
            try
            {
                Debug.WriteLine("开始执行格式化插入内容预览...");

                // 获取参数
                string targetHeading = GetParameterValue<string>(parameters, "target_heading", "");
                string content = GetParameterValue<string>(parameters, "content", "");
                string formatType = GetParameterValue<string>(parameters, "format_type", "paragraph");
                int indentLevel = GetParameterValue<int>(parameters, "indent_level", 0);
                bool addSpacing = GetParameterValue<bool>(parameters, "add_spacing", false); // 默认不添加额外间距
                string insertPositionRaw = GetParameterValue<string>(parameters, "insert_position", _defaultInsertPosition);
                string insertPosition = NormalizeInsertPosition(!string.IsNullOrEmpty(insertPositionRaw) ? insertPositionRaw : _defaultInsertPosition);
                NotifyProgress($"formatted_insert_content: 插入位置参数 raw='{insertPositionRaw}', 默认='{_defaultInsertPosition}', 使用='{insertPosition}'");
                bool previewOnly = GetParameterValue<bool>(parameters, "preview_only", true); // 默认只预览

                // 内容不能为空，但标题可以为空（表示在当前光标位置插入）
                if (string.IsNullOrEmpty(content))
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "插入内容不能为空" });
                }

                // 如果只是预览，返回格式化的预览内容
                if (previewOnly)
                {
                    string previewContent = GeneratePreviewContent(content, formatType, indentLevel, addSpacing);

                    // 根据是否有标题生成不同的预览信息
                    string previewMessage;
                    if (string.IsNullOrEmpty(targetHeading))
                    {
                        previewMessage = $"预览：将在当前光标位置插入{GetFormatTypeName(formatType)}格式的内容";
                    }
                    else
                    {
                        previewMessage = $"预览：将在标题 '{targetHeading}' 下方插入{GetFormatTypeName(formatType)}格式的内容";
                    }

                    return JsonConvert.SerializeObject(new
                    {
                        success = true,
                        action_type = "insert_content",
                        preview_mode = true,
                        target_heading = targetHeading ?? "",
                        format_type = formatType,
                        indent_level = indentLevel,
                        add_spacing = addSpacing,
                        insert_position = insertPosition,
                        original_content = content,
                        preview_content = previewContent,
                        message = previewMessage
                    });
                }

                // 实际执行插入
                return ExecuteFormattedInsert(targetHeading, content, formatType, indentLevel, addSpacing, insertPosition);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"格式化插入内容时出错: {ex.Message}");
                return JsonConvert.SerializeObject(new { success = false, message = $"插入失败: {ex.Message}" });
            }
        }

        // 生成预览内容
        private static string GeneratePreviewContent(string content, string formatType, int indentLevel, bool addSpacing)
        {
            string indentPrefix = new string(' ', indentLevel * 4); // 每级缩进4个空格用于预览

            switch (formatType.ToLower())
            {
                case "list":
                    string[] lines = content.Split('\n');
                    var listItems = lines.Where(line => !string.IsNullOrWhiteSpace(line))
                                         .Select(line => $"{indentPrefix}• {line.Trim()}");
                    return string.Join("\n", listItems); // 不添加前后间距

                case "table":
                    return $"{indentPrefix}[表格预览]\n{indentPrefix}{content.Replace("\n", $"\n{indentPrefix}")}";

                case "emphasis":
                    return $"{indentPrefix}**{content}**";

                default: // paragraph
                    return $"{indentPrefix}{content}";
            }
        }

        // 获取格式类型中文名称
        private static string GetFormatTypeName(string formatType)
        {
            switch (formatType.ToLower())
            {
                case "list": return "列表";
                case "table": return "表格";
                case "emphasis": return "强调";
                default: return "段落";
            }
        }

        // 执行实际的格式化插入
        private static string ExecuteFormattedInsert(string targetHeading, string content, string formatType, int indentLevel, bool addSpacing, string insertPosition = "end")
        {
            try
            {
                // 获取Word应用程序实例
                dynamic wordApp = _markdownToWord.GetWordApplication();
                if (wordApp == null)
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "无法连接到Word应用程序" });
                }

                dynamic doc = _markdownToWord.GetActiveDocument(wordApp);
                if (doc == null)
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "无法获取活动文档" });
                }

                dynamic titleRange;

                // 处理有标题和无标题的情况
                if (string.IsNullOrEmpty(targetHeading))
                {
                    // 没有标题时，在当前光标位置插入
                    Debug.WriteLine("没有指定标题，将在当前光标位置插入内容");
                    titleRange = wordApp.Selection.Range;
                }
                else
                {
                    // 使用基于标题样式的精准查找，避免匹配正文中的同名文本
                    var headingResult = FindHeadingEfficiently(doc, targetHeading, "keywords");
                    if (!headingResult.found)
                    {
                        return JsonConvert.SerializeObject(new
                        {
                            success = false,
                            message = $"未找到标题: {targetHeading}。请检查标题是否存在，或使用标题中的关键词。",
                            suggestions = headingResult.suggestions ?? new string[0]
                        });
                    }

                    dynamic headingRange = headingResult.headingInfo.range;

                    // 读取父标题层级（优先OutlineLevel，失败回退样式解析）
                    int parentLevel = 1;
                    try
                    {
                        dynamic headingPara = headingRange.Paragraphs[1];
                        try
                        {
                            var ol = headingPara.OutlineLevel;
                            if (ol != null && ol is int) parentLevel = (int)ol;
                        }
                        catch { }
                        if (parentLevel < 1 || parentLevel > 9)
                        {
                            try
                            {
                                string styleName = "";
                                try { styleName = headingPara.Style?.NameLocal ?? headingPara.Style?.Name ?? headingPara.Style.ToString(); } catch { }
                                parentLevel = ExtractHeadingLevel(styleName);
                                if (parentLevel < 1 || parentLevel > 9) parentLevel = 1;
                            }
                            catch { parentLevel = 1; }
                        }
                    }
                    catch { parentLevel = 1; }

                    // 在插入前，基于父标题层级清理/调整生成内容中的标题：
                    // - 移除与父标题同名且级别任意的重复标题（常见为再次输出父级标题）
                    // - 若出现级别<=父级的标题，则降级为父级+1
                    try
                    {
                        if (!string.IsNullOrEmpty(content))
                        {
                            content = SanitizeHeadingsForInsertion(content, targetHeading, parentLevel);
                        }
                    }
                    catch (Exception sanitizeEx)
                    {
                        Debug.WriteLine($"清理生成内容中的标题时出错: {sanitizeEx.Message}");
                    }

                    // 根据插入位置参数选择插入策略
                    if (insertPosition == "front")
                    {
                        // 插入到标题紧接着的位置
                        titleRange = headingRange.Duplicate;
                        titleRange.Collapse(0); // 移动到标题末尾
                        Debug.WriteLine($"插入到标题前端位置: {titleRange.Start}");
                    }
                    else
                    {
                        // 默认：智能定位插入位置，在标题下已有内容的末尾插入
                        titleRange = FindBestInsertionPoint(doc, headingRange);
                        Debug.WriteLine($"智能定位到插入位置: {titleRange.Start}");
                    }
                }

                // 使用书签标记插入范围起点，确保后续清理只作用于本次插入内容
                string insertStartBookmark = "OWC_START_" + Guid.NewGuid().ToString("N");
                try { doc.Bookmarks.Add(insertStartBookmark, doc.Range(titleRange.Start, titleRange.Start)); } catch { }

                // 只有在找到标题的情况下才检查样式和换行
                if (!string.IsNullOrEmpty(targetHeading))
                {
                    // 获取当前段落信息
                    dynamic currentPara = titleRange.Paragraphs[1];
                    string currentStyle = "";
                    try
                    {
                        try { currentStyle = currentPara.Style?.NameLocal ?? ""; } catch { currentStyle = currentPara.Style.ToString(); }
                        Debug.WriteLine($"当前段落样式: {currentStyle}");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"获取段落样式时出错: {ex.Message}");
                    }

                    // 检查是否是标题样式
                    bool isHeadingStyle = IsHeadingStyle(currentStyle);
                    Debug.WriteLine($"是否为标题样式: {isHeadingStyle}");

                    // 仅在选择 front（标题后）插入时才需要从标题段创建新正文段落
                    if (insertPosition == "front")
                    {
                        if (isHeadingStyle)
                        {
                            // 如果当前在标题段落中，仅添加一次换行，开始新段落
                            Debug.WriteLine("在标题段落中，添加一次换行并创建新段落");
                            titleRange.InsertAfter("\r\n");
                            titleRange.Collapse(0); // 移动到新行的开始

                            // 确保新段落使用正文样式
                            try
                            {
                                dynamic newPara = titleRange.Paragraphs[1];
                                newPara.Style = "Normal"; // 设置为正文样式
                                Debug.WriteLine("新段落样式已设置为Normal");
                            }
                            catch (Exception styleEx)
                            {
                                Debug.WriteLine($"设置段落样式时出错: {styleEx.Message}");
                            }
                        }
                        else
                        {
                            // front 模式下，如果光标未落在标题段，尽量避免错位，保持当前位置
                            Debug.WriteLine("front 模式但当前位置非标题段，保持当前位置不前移");
                        }
                    }
                    // insertPosition == "end" 时，titleRange 已经是标题下已有内容的末尾，不做段落前移，避免越过到下一个标题

                    // 插入前间距：仅在非常必要时添加（默认情况下不添加）
                    if (addSpacing)
                    {
                        Debug.WriteLine("用户明确要求添加前置间距");
                        titleRange.InsertAfter("\r\n");
                        titleRange.Collapse(0);
                    }
                }
                else
                {
                    // 没有指定标题，在当前光标位置插入；仅在需要时添加前置换行
                    Debug.WriteLine("在当前光标位置插入内容");
                    if (addSpacing && titleRange.Start > 0)
                    {
                        titleRange.InsertBefore("\r\n");
                        titleRange.Collapse(0);
                    }
                }

                // 根据格式类型插入内容
                switch (formatType.ToLower())
                {
                    case "list":
                        InsertListContent(titleRange, content, indentLevel);
                        break;
                    case "table":
                        InsertTableContent(titleRange, content, indentLevel);
                        break;
                    case "emphasis":
                        InsertEmphasisContent(titleRange, content, indentLevel);
                        break;
                    default: // paragraph
                        // 智能检测：若段落中包含表格结构，则拆分为“前文/表格/后文”分别插入
                        if (ContainsTableStructure(content))
                        {
                            Debug.WriteLine("检测到段落中包含表格结构，执行混合内容插入");
                            InsertMixedContent(titleRange, content, indentLevel);
                        }
                        else
                        {
                            InsertParagraphContent(titleRange, content, indentLevel);
                        }
                        break;
                }

                if (addSpacing)
                {
                    // 插入后间距：仅在用户明确要求时添加
                    Debug.WriteLine("用户明确要求添加后置间距");
                    titleRange.InsertAfter("\r\n");
                }

                // 清理：仅清理书签之间的本次插入内容范围
                try
                {
                    // 在当前光标处打上结束书签
                    string insertEndBookmark = "OWC_END_" + Guid.NewGuid().ToString("N");
                    try { doc.Bookmarks.Add(insertEndBookmark, doc.Range(wordApp.Selection.Range.End, wordApp.Selection.Range.End)); } catch { }

                    // 精确获取起止范围
                    dynamic startRange = null;
                    dynamic endRange = null;

                    try { startRange = doc.Bookmarks[insertStartBookmark]?.Range; } catch { }
                    try { endRange = doc.Bookmarks[insertEndBookmark]?.Range; } catch { }

                    if (startRange != null && endRange != null)
                    {
                        try
                        {
                            int rangeLength = endRange.End - startRange.Start;
                            Debug.WriteLine($"插入范围长度: {rangeLength} 字符");

                            // 只对合理大小的范围进行清理，避免处理过大的内容
                            if (_enableRangeNormalization && rangeLength > 0 && rangeLength < 5000) // 限制范围大小
                            {
                                dynamic insertedRange = doc.Range(startRange.Start, endRange.End);
                                NormalizeInsertedRange(insertedRange, keepOneTrailingEmpty: addSpacing);
                            }
                            else
                            {
                                if (!_enableRangeNormalization)
                                {
                                    Debug.WriteLine("范围清理功能已禁用");
                                }
                                else
                                {
                                    Debug.WriteLine($"跳过范围清理：范围过大 ({rangeLength} 字符) 或无效");
                                }
                            }
                        }
                        catch (Exception rangeEx)
                        {
                            Debug.WriteLine($"处理插入范围时出错: {rangeEx.Message}");
                        }
                    }
                    else
                    {
                        Debug.WriteLine("无法获取插入范围书签，跳过清理");
                    }

                    // 清理书签
                    try { doc.Bookmarks[insertStartBookmark]?.Delete(); } catch { }
                    try { doc.Bookmarks[insertEndBookmark]?.Delete(); } catch { }
                }
                catch (Exception cleanupEx)
                {
                    Debug.WriteLine($"清理插入范围时出错: {cleanupEx.Message}");
                }

                return JsonConvert.SerializeObject(new
                {
                    success = true,
                    message = $"成功在标题 '{targetHeading}' 下方插入{GetFormatTypeName(formatType)}格式的内容",
                    heading = targetHeading,
                    format_type = formatType,
                    content_length = content.Length
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"执行格式化插入时出错: {ex.Message}");
                return JsonConvert.SerializeObject(new { success = false, message = $"插入失败: {ex.Message}" });
            }
        }

        // 清理/调整生成内容中的Markdown标题：
        // 1) 如果首个（或任意）标题文本与父标题相同，则删除该行（避免重复父标题）
        // 2) 若发现标题级别<=父标题级别，则降级为 父标题级别+1
        private static string SanitizeHeadingsForInsertion(string content, string targetHeading, int parentLevel)
        {
            if (string.IsNullOrWhiteSpace(content)) return content;
            string normalizedTarget = NormalizeText(targetHeading ?? "");

            var lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            var sb = new System.Text.StringBuilder();
            bool seenFirstNonEmpty = false;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                string trimmed = (line ?? "").Trim();

                if (!seenFirstNonEmpty && string.IsNullOrWhiteSpace(trimmed))
                {
                    sb.AppendLine(line);
                    continue;
                }
                if (!seenFirstNonEmpty) seenFirstNonEmpty = true;

                var m = System.Text.RegularExpressions.Regex.Match(trimmed, @"^(#{1,6})\s*(.+)$");
                if (m.Success)
                {
                    int hLevel = m.Groups[1].Value.Length;
                    string hText = m.Groups[2].Value.Trim();
                    string normalizedHeader = NormalizeText(hText);

                    // 规则1：与父标题同名 → 删除该行（避免重复输出父标题）
                    if (!string.IsNullOrEmpty(normalizedTarget) && normalizedHeader == normalizedTarget)
                    {
                        // 跳过本行；若下一行是单独的空行也一并跳过，避免多余空白
                        if (i + 1 < lines.Length && string.IsNullOrWhiteSpace(lines[i + 1]))
                        {
                            i++;
                        }
                        continue;
                    }

                    // 规则2：级别不合法（<=父级） → 强制降为父级+1
                    if (hLevel <= parentLevel)
                    {
                        int newLevel = Math.Min(parentLevel + 1, 6);
                        string newPrefix = new string('#', newLevel);
                        // 用新的 # 前缀替换
                        line = System.Text.RegularExpressions.Regex.Replace(line, @"^(#{1,6})", newPrefix);
                    }
                }

                sb.AppendLine(line);
            }

            return sb.ToString().TrimEnd();
        }

        // 清除范围内的软回车（手动换行符）
        private static void RemoveManualLineBreaksInRange(dynamic range)
        {
            try
            {
                if (range == null) return;

                Debug.WriteLine("开始清除软回车（手动换行符）");

                // 保存当前选择位置
                dynamic wordApp = range.Application;
                dynamic originalSelection = wordApp.Selection.Range;
                int originalStart = originalSelection.Start;
                int originalEnd = originalSelection.End;

                // 创建Find对象，在指定范围内查找
                dynamic find = range.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();

                // 设置查找软回车（^l 在Word中表示手动换行符）
                find.Text = "^l";
                find.Replacement.Text = " "; // 替换为空格，避免词语连在一起
                find.Forward = true;
                find.Wrap = 0; // wdFindStop = 0, 不循环查找，只在range内查找
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;

                // 执行全部替换（wdReplaceAll = 2）
                int replaceCount = 0;
                bool found = true;
                
                // 循环替换，直到没有找到为止
                while (found)
                {
                    try
                    {
                        found = find.Execute(
                            FindText: Type.Missing,
                            MatchCase: Type.Missing,
                            MatchWholeWord: Type.Missing,
                            MatchWildcards: Type.Missing,
                            MatchSoundsLike: Type.Missing,
                            MatchAllWordForms: Type.Missing,
                            Forward: Type.Missing,
                            Wrap: Type.Missing,
                            Format: Type.Missing,
                            ReplaceWith: Type.Missing,
                            Replace: 2 // wdReplaceAll = 2
                        );
                        
                        if (found)
                        {
                            replaceCount++;
                            if (replaceCount > 1000) // 防止无限循环
                            {
                                Debug.WriteLine("软回车清除次数超过1000次，可能存在异常，停止清除");
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"清除软回车时出错: {ex.Message}");
                        break;
                    }
                }

                if (replaceCount > 0)
                {
                    Debug.WriteLine($"已清除 {replaceCount} 个软回车");
                }
                else
                {
                    Debug.WriteLine("未发现软回车");
                }

                // 清理连续的空格（软回车替换后可能产生）
                try
                {
                    find.ClearFormatting();
                    find.Replacement.ClearFormatting();
                    find.Text = "  "; // 两个空格
                    find.Replacement.Text = " "; // 替换为一个空格
                    find.Forward = true;
                    find.Wrap = 0;
                    
                    // 循环替换连续空格
                    while (find.Execute(
                        FindText: Type.Missing,
                        MatchCase: Type.Missing,
                        MatchWholeWord: Type.Missing,
                        MatchWildcards: Type.Missing,
                        MatchSoundsLike: Type.Missing,
                        MatchAllWordForms: Type.Missing,
                        Forward: Type.Missing,
                        Wrap: Type.Missing,
                        Format: Type.Missing,
                        ReplaceWith: Type.Missing,
                        Replace: 2
                    )) { }
                }
                catch { }

                // 恢复原始选择位置
                try
                {
                    wordApp.Selection.SetRange(originalStart, originalEnd);
                }
                catch { }

                Debug.WriteLine("软回车清除完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"RemoveManualLineBreaksInRange 失败: {ex.Message}");
            }
        }

        // 规范化插入后的范围：
        // - 统一段前后距为0，行距为单倍
        // - 移除多余空段落（只保留最后一个空段，若 keepOneTrailingEmpty 为 true）
        private static void NormalizeInsertedRange(dynamic range, bool keepOneTrailingEmpty)
        {
            try
            {
                if (range == null) return;

                // 首先清除所有软回车（手动换行符）
                RemoveManualLineBreaksInRange(range);

                dynamic paragraphs = range.Paragraphs;
                if (paragraphs == null) return;

                int count = 0;
                try
                {
                    count = paragraphs.Count;
                    if (count <= 0) return;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"无法获取段落数量: {ex.Message}");
                    return;
                }

                Debug.WriteLine($"开始规范化 {count} 个段落");

                // 统一段落格式 - 更安全的方式
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic p = paragraphs[i];
                        if (p == null) continue;

                        // 安全地设置段落格式
                        try { p.SpaceBefore = 0; } catch { }
                        try { p.SpaceAfter = 0; } catch { }
                        try { p.LineSpacingRule = 0; } catch { } // 0 = wdLineSpaceSingle
                        try { p.NoSpaceBetweenParagraphsOfSameStyle = 1; } catch { }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"设置段落 {i} 格式时出错: {ex.Message}");
                    }
                }

                // 移除多余的空段落 - 更安全的方式
                for (int i = count; i >= 1; i--)
                {
                    try
                    {
                        dynamic p = paragraphs[i];
                        if (p == null) continue;

                        string text = "";
                        try
                        {
                            dynamic pRange = p.Range;
                            if (pRange != null)
                            {
                                text = pRange.Text as string ?? "";
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"获取段落 {i} 文本时出错: {ex.Message}");
                            continue;
                        }

                        string trimmed = text.Replace("\r", "").Replace("\n", "").Trim();
                        bool isEmpty = string.IsNullOrEmpty(trimmed);

                        bool isLast = (i == count);
                        if (isEmpty)
                        {
                            // 如果是最后一个空段，根据参数决定是否保留；其他空段全部删除
                            if (!(isLast && keepOneTrailingEmpty))
                            {
                                try
                                {
                                    dynamic pRange = p.Range;
                                    if (pRange != null)
                                    {
                                        pRange.Delete();
                                        Debug.WriteLine($"删除空段落 {i}");
                                    }
                                }
                                catch (Exception delEx)
                                {
                                    Debug.WriteLine($"删除段落 {i} 时出错: {delEx.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"处理段落 {i} 时出错: {ex.Message}");
                    }
                }

                Debug.WriteLine("段落规范化完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"NormalizeInsertedRange 失败: {ex.Message}");
            }
        }

        // 插入段落内容 - 使用MarkdownToWord进行格式化
        private static void InsertParagraphContent(dynamic range, string content, int indentLevel)
        {
            try
            {
                Debug.WriteLine($"插入段落内容: {content.Substring(0, Math.Min(100, content.Length))}...");

                // 移动光标到插入位置
                dynamic wordApp = range.Application;
                dynamic selection = wordApp.Selection;
                selection.SetRange(range.Start, range.Start);

                // 检测内容类型并使用适当的插入方法
                if (IsMarkdownContent(content))
                {
                    Debug.WriteLine("检测到Markdown内容，使用HTML插入方法");
                    // 如果是Markdown内容，转换为HTML后插入
                    string htmlContent = ConvertMarkdownToHtml(content);
                    _markdownToWord.InsertHtmlContentDirect(htmlContent);
                }
                else
                {
                    Debug.WriteLine("使用文本插入方法");
                    // 普通文本直接插入
                    _markdownToWord.InsertText(content);
                }

                // 设置缩进
                if (indentLevel > 0)
                {
                    try
                    {
                        selection.ParagraphFormat.LeftIndent = indentLevel * 18; // 每级缩进18磅
                        Debug.WriteLine($"已设置段落缩进: {indentLevel * 18}磅");
                    }
                    catch (Exception indentEx)
                    {
                        Debug.WriteLine($"设置缩进时出错: {indentEx.Message}");
                    }
                }

                Debug.WriteLine("段落内容插入完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入段落内容时出错: {ex.Message}");
                // 回退到简单插入
                range.InsertAfter(content);
                range.Collapse(0);
            }
        }

        // 插入列表内容 - 简化版本，减少多余换行
        private static void InsertListContent(dynamic range, string content, int indentLevel)
        {
            try
            {
                Debug.WriteLine($"插入列表内容: {content.Substring(0, Math.Min(100, content.Length))}...");

                // 移动光标到插入位置
                dynamic wordApp = range.Application;
                dynamic selection = wordApp.Selection;
                selection.SetRange(range.Start, range.Start);

                // 分析内容，确定是有序列表还是无序列表
                string[] lines = content.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                bool isOrderedList = false;

                // 检查是否为有序列表（兼容形如 1) 1. 等）
                foreach (string line in lines)
                {
                    string trimmedLine = line.TrimStart();
                    if (System.Text.RegularExpressions.Regex.IsMatch(trimmedLine, @"^\d+[\.\)]\s"))
                    {
                        isOrderedList = true;
                        break;
                    }
                }

                Debug.WriteLine($"列表类型: {(isOrderedList ? "有序列表" : "无序列表")}");

                // 先插入所有文本内容，然后统一应用列表格式
                var nonEmptyLines = lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
                for (int i = 0; i < nonEmptyLines.Length; i++)
                {
                    string line = nonEmptyLines[i].Trim();

                    // 清理已有的列表标记（兼容多种项目符号，包括 • 等）
                    line = System.Text.RegularExpressions.Regex.Replace(line, @"^\s*(?:[-\*\+•●◦▪▫·‣⁃–—])\s*", "");
                    line = System.Text.RegularExpressions.Regex.Replace(line, @"^\s*\d+[\.\)]\s*", "");

                    // 插入文本
                    selection.TypeText(line);

                    // 只在非最后一项时添加段落分隔
                    if (i < nonEmptyLines.Length - 1)
                    {
                        selection.TypeText("\r");
                    }
                }

                // 选择刚插入的所有内容
                dynamic insertedRange = range.Document.Range(range.Start, selection.Range.End);
                selection.SetRange(insertedRange.Start, insertedRange.End);

                // 统一应用列表格式
                try
                {
                    if (isOrderedList)
                    {
                        selection.Range.ListFormat.ApplyNumberDefault();
                    }
                    else
                    {
                        selection.Range.ListFormat.ApplyBulletDefault();
                    }

                    // 设置缩进级别
                    if (indentLevel > 0)
                    {
                        for (int j = 0; j < indentLevel; j++)
                        {
                            selection.Range.ListFormat.ListIndent();
                        }
                    }

                    Debug.WriteLine($"已统一应用列表格式，项目数: {nonEmptyLines.Length}");
                }
                catch (Exception formatEx)
                {
                    Debug.WriteLine($"应用列表格式时出错: {formatEx.Message}");
                }

                // 将光标移动到插入内容的末尾
                selection.Collapse(0); // wdCollapseEnd

                Debug.WriteLine("列表内容插入完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入列表内容时出错: {ex.Message}");
                // 回退到最简单的插入方法
                try
                {
                    string[] lines = content.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    dynamic wordApp = range.Application;
                    dynamic selection = wordApp.Selection;
                    selection.SetRange(range.Start, range.Start);

                    var nonEmptyLines = lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
                    for (int i = 0; i < nonEmptyLines.Length; i++)
                    {
                        string line = nonEmptyLines[i].Trim();
                        string cleanLine = System.Text.RegularExpressions.Regex.Replace(line, @"^\s*(?:[-\*\+•●◦▪▫·‣⁃–—])\s*", "");
                        cleanLine = System.Text.RegularExpressions.Regex.Replace(cleanLine, @"^\s*\d+[\.\)]\s*", "");

                        selection.TypeText("• " + cleanLine);

                        // 只在非最后一项时添加换行
                        if (i < nonEmptyLines.Length - 1)
                        {
                            selection.TypeText("\r");
                        }
                    }
                    Debug.WriteLine("使用回退方法插入列表");
                }
                catch (Exception fallbackEx)
                {
                    Debug.WriteLine($"回退插入方法也失败: {fallbackEx.Message}");
                    range.InsertAfter(content);
                }
            }
        }

        // 插入强调内容 - 使用MarkdownToWord进行格式化
        private static void InsertEmphasisContent(dynamic range, string content, int indentLevel)
        {
            try
            {
                Debug.WriteLine($"插入强调内容: {content.Substring(0, Math.Min(100, content.Length))}...");

                // 移动光标到插入位置
                dynamic wordApp = range.Application;
                dynamic selection = wordApp.Selection;
                selection.SetRange(range.Start, range.Start);

                // 将内容包装为Markdown粗体格式
                string emphasizedContent = $"**{content}**";
                Debug.WriteLine($"强调内容包装为: {emphasizedContent}");

                // 转换为HTML并插入
                string htmlContent = ConvertMarkdownToHtml(emphasizedContent);
                _markdownToWord.InsertHtmlContentDirect(htmlContent);

                // 设置缩进
                if (indentLevel > 0)
                {
                    try
                    {
                        selection.ParagraphFormat.LeftIndent = indentLevel * 18;
                        Debug.WriteLine($"已设置强调内容缩进: {indentLevel * 18}磅");
                    }
                    catch (Exception indentEx)
                    {
                        Debug.WriteLine($"设置强调内容缩进时出错: {indentEx.Message}");
                    }
                }

                Debug.WriteLine("强调内容插入完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入强调内容时出错: {ex.Message}");
                // 回退到简单插入
                range.InsertAfter(content);
                range.Collapse(0);

                try
                {
                    dynamic lastPara = range.Paragraphs[range.Paragraphs.Count];
                    lastPara.Style = "Normal";
                    lastPara.Range.Font.Bold = true;
                    if (indentLevel > 0)
                    {
                        lastPara.LeftIndent = indentLevel * 18;
                    }
                }
                catch (Exception styleEx)
                {
                    Debug.WriteLine($"回退样式设置时出错: {styleEx.Message}");
                }
            }
        }

        // 插入表格内容（修复版：解决列对齐和粗体格式问题）
        private static void InsertTableContent(dynamic range, string content, int indentLevel)
        {
            // 清理可能的预览前缀
            string cleanContent = content;
            if (cleanContent.StartsWith("[表格内容]"))
            {
                cleanContent = cleanContent.Substring(6).TrimStart('\n', '\r', ' ');
                Debug.WriteLine("检测到并移除了预览前缀 '[表格内容]'");
            }
            else if (cleanContent.StartsWith("[表格预览]"))
            {
                cleanContent = cleanContent.Substring(6).TrimStart('\n', '\r', ' ');
                Debug.WriteLine("检测到并移除了预览前缀 '[表格预览]'");
            }

            // 若前面存在非表格的描述文本（例如“介绍性段落 + |... --- ...|”），改为走混合内容流程，避免把描述行当成表头
            try
            {
                string trimmed = cleanContent.TrimStart();
                if (!trimmed.StartsWith("|") && cleanContent.Contains("|") && cleanContent.Contains("---"))
                {
                    Debug.WriteLine("检测到表格前存在描述文本，改用 InsertMixedContent 处理");
                    InsertMixedContent(range, cleanContent, indentLevel);
                    return;
                }
            }
            catch { }

            string[] lines = cleanContent.Split('\n');
            if (lines.Length < 2) return;

            try
            {
                Debug.WriteLine($"原始表格内容: {content}");
                Debug.WriteLine($"清理后表格内容: {cleanContent}");

                // 计算列数（从第一行）- 修复列对齐问题
                string[] headers = SplitTableRow(lines[0]);

                int cols = headers.Length;
                if (cols == 0) return;

                Debug.WriteLine($"表头解析结果: [{string.Join(", ", headers)}], 列数: {cols}");

                // 处理数据行：严格只采集真正的表格行，遇到第一条非表格行立即停止
                var dataLinesList = new List<string>();
                bool seenSeparator = false; // 确保分隔线之后才开始采集
                for (int i = 1; i < lines.Length; i++)
                {
                    var raw = lines[i];
                    var trimmed = (raw ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(trimmed))
                        continue;

                    if (!seenSeparator)
                    {
                        if (IsMarkdownTableSeparator(trimmed))
                        {
                            seenSeparator = true;
                            continue; // 分隔线本身不计入数据行
                        }
                        // 在遇到分隔线前的内容全部忽略（防御性）
                        continue;
                    }

                    // 分隔线之后：只接受看起来像表格行的行
                    if (IsLikelyTableRow(trimmed, cols))
                    {
                        dataLinesList.Add(trimmed);
                        if (dataLinesList.Count >= 20) break; // 上限
                    }
                    else
                    {
                        // 如果这一行包含“|”，尝试截断到最后一个“|”作为表格行，余下算正文
                        if (trimmed.StartsWith("|") && trimmed.Contains("|"))
                        {
                            int lastPipe = trimmed.LastIndexOf('|');
                            if (lastPipe > 0)
                            {
                                string candidate = trimmed.Substring(0, lastPipe + 1).TrimEnd();
                                if (IsLikelyTableRow(candidate, cols))
                                {
                                    dataLinesList.Add(candidate);
                                }
                            }
                        }
                        break; // 一旦遇到非表格行（或截断后），立即停止
                    }
                }
                var dataLines = dataLinesList.ToArray();

                int rows = dataLines.Length + 1; // +1 for header row

                Debug.WriteLine($"开始插入表格: {rows} 行 x {cols} 列");

                // 插入表格
                dynamic table = range.Tables.Add(range, rows, cols);

                // 设置表格样式
                try
                {
                    table.Borders.InsideLineStyle = 1; // wdLineStyleSingle
                    table.Borders.OutsideLineStyle = 1;
                    table.Borders.InsideLineWidth = 2; // wdLineWidth025pt
                    table.Borders.OutsideLineWidth = 2;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"设置表格边框时出错: {ex.Message}");
                }

                // 设置表头（第一行）
                for (int j = 0; j < cols; j++)
                {
                    try
                    {
                        var cell = table.Cell(1, j + 1);

                        // 处理表头的Markdown格式（如粗体）
                        string headerText = ProcessCellMarkdown(headers[j]);
                        SetCellContent(cell, headerText, true); // 表头强制加粗

                        // 设置表头背景和对齐
                        cell.Shading.BackgroundPatternColor = 15658734; // 浅灰色背景
                        cell.Range.ParagraphFormat.Alignment = 1; // 居中对齐

                        Debug.WriteLine($"表头 [{j + 1}]: {headerText}");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"设置表头格式时出错: {ex.Message}");
                    }
                }

                // 填充数据行
                for (int i = 0; i < dataLines.Length; i++)
                {
                    string[] cells = SplitTableRow(dataLines[i]);
                    Debug.WriteLine($"数据行 [{i + 1}] 解析结果: [{string.Join(", ", cells)}]");

                    for (int j = 0; j < cols; j++)
                    {
                        try
                        {
                            var cell = table.Cell(i + 2, j + 1); // +2 because row 1 is header

                            // 获取单元格内容，如果超出范围则为空
                            string cellValue = j < cells.Length ? cells[j] : "";

                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                // 处理单元格的Markdown格式
                                string processedText = ProcessCellMarkdown(cellValue);
                                SetCellContent(cell, processedText, false); // 数据单元格不强制加粗

                                Debug.WriteLine($"数据单元格 [{i + 2},{j + 1}]: {processedText}");
                            }

                            // 设置数据单元格对齐
                            cell.Range.ParagraphFormat.Alignment = 0; // 左对齐
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"填充单元格 [{i + 2},{j + 1}] 时出错: {ex.Message}");
                        }
                    }
                }

                // 自动调整表格
                try
                {
                    table.AutoFitBehavior(1); // wdAutoFitContent
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"自动调整表格时出错: {ex.Message}");
                }

                Debug.WriteLine($"表格插入成功: {rows} 行 x {cols} 列");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入表格时出错: {ex.Message}");
                Debug.WriteLine($"错误详情: {ex}");
                // 回退到段落格式
                InsertParagraphContent(range, content, 0);
            }
        }

        // 辅助方法：智能分割表格行（修复列对齐问题）
        private static string[] SplitTableRow(string row)
        {
            if (string.IsNullOrWhiteSpace(row)) return new string[0];

            // 移除行首尾的 | 符号
            string cleanRow = row.Trim();
            if (cleanRow.StartsWith("|")) cleanRow = cleanRow.Substring(1);
            if (cleanRow.EndsWith("|")) cleanRow = cleanRow.Substring(0, cleanRow.Length - 1);

            // 按 | 分割并清理每个单元格
            string[] cells = cleanRow.Split('|')
                .Select(cell => cell.Trim())
                .ToArray();

            return cells;
        }

        // 辅助方法：判断是否为Markdown表格分隔符行
        private static bool IsMarkdownTableSeparator(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;

            string trimmed = line.Trim();
            // 检查是否只包含 |、-、: 和空格字符
            return trimmed.All(c => c == '|' || c == '-' || c == ':' || c == ' ');
        }

        // 判断是否可能是表格数据行：以 | 开头和结尾，且单元格数与列数匹配
        private static bool IsLikelyTableRow(string line, int expectedCols)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;
            var s = line.Trim();
            if (!s.StartsWith("|") || !s.EndsWith("|")) return false;
            // 分割并统计列（去除空白）
            var inner = s.Substring(1, s.Length - 2);
            var cells = inner.Split('|').Select(c => c.Trim()).ToArray();
            if (cells.Length != expectedCols) return false;
            // 粗略过滤明显的非表格内容（标题、列表等）
            foreach (var c in cells)
            {
                var t = c.TrimStart();
                if (t.StartsWith("##") || t.StartsWith("###") || (t.StartsWith("-") && !t.StartsWith("---")) || (t.StartsWith("*") && !t.StartsWith("**")))
                {
                    return false;
                }
            }
            return true;
        }

        // 辅助方法：处理单元格中的Markdown格式
        private static string ProcessCellMarkdown(string cellContent)
        {
            if (string.IsNullOrWhiteSpace(cellContent)) return "";

            string processed = cellContent.Trim();

            // 移除Markdown粗体语法，但保留文本内容用于后续格式化
            // 这里只是清理文本，实际的粗体格式在SetCellContent中应用
            processed = processed.Replace("**", "");
            processed = processed.Replace("__", "");

            return processed;
        }

        // 辅助方法：设置单元格内容和格式
        private static void SetCellContent(dynamic cell, string content, bool isHeader)
        {
            try
            {
                // 检查原始内容是否包含粗体标记（在清理之前检查）
                bool shouldBeBold = isHeader || content.Contains("**") || content.Contains("__");

                // 设置文本内容（清理Markdown语法）
                string cleanText = ProcessCellMarkdown(content);
                cell.Range.Text = cleanText;

                // 应用格式
                if (shouldBeBold)
                {
                    cell.Range.Font.Bold = true;
                    Debug.WriteLine($"单元格设为粗体: {cleanText} (原始: {content})");
                }
                else
                {
                    cell.Range.Font.Bold = false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"设置单元格内容时出错: {ex.Message}");
                // 回退到简单文本设置
                try
                {
                    cell.Range.Text = ProcessCellMarkdown(content);
                }
                catch (Exception fallbackEx)
                {
                    Debug.WriteLine($"回退设置也失败: {fallbackEx.Message}");
                }
            }
        }

        // 文字样式修改实现
        private static string ModifyTextStyle(Dictionary<string, object> parameters)
        {
            try
            {
                Debug.WriteLine("开始执行文字样式修改预览...");

                // 获取参数
                string textToFind = GetParameterValue<string>(parameters, "text_to_find", "");
                bool previewOnly = GetParameterValue<bool>(parameters, "preview_only", true); // 默认只预览

                if (string.IsNullOrEmpty(textToFind))
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "要查找的文本不能为空" });
                }

                // 如果只是预览，返回样式预览信息
                if (previewOnly)
                {
                    List<string> previewStyles = GenerateStylePreview(parameters);
                    return JsonConvert.SerializeObject(new
                    {
                        success = true,
                        action_type = "modify_style",
                        preview_mode = true,
                        text_to_find = textToFind,
                        preview_styles = previewStyles,
                        style_parameters = parameters,
                        message = $"预览：将为文本 '{textToFind}' 应用以下样式"
                    });
                }

                // 实际执行样式修改
                return ExecuteStyleModification(parameters);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"修改文字样式时出错: {ex.Message}");
                return JsonConvert.SerializeObject(new { success = false, message = $"样式修改失败: {ex.Message}" });
            }
        }

        // 生成样式预览信息
        private static List<string> GenerateStylePreview(Dictionary<string, object> parameters)
        {
            List<string> previewStyles = new List<string>();

            if (parameters.ContainsKey("font_size"))
            {
                int fontSize = GetParameterValue<int>(parameters, "font_size", 0);
                if (fontSize >= 8 && fontSize <= 72)
                {
                    previewStyles.Add($"🔤 字体大小: {fontSize}磅");
                }
            }

            if (parameters.ContainsKey("font_bold"))
            {
                bool bold = GetParameterValue<bool>(parameters, "font_bold", false);
                previewStyles.Add($"🔲 粗体: {(bold ? "是" : "否")}");
            }

            if (parameters.ContainsKey("font_italic"))
            {
                bool italic = GetParameterValue<bool>(parameters, "font_italic", false);
                previewStyles.Add($"📐 斜体: {(italic ? "是" : "否")}");
            }

            if (parameters.ContainsKey("font_color"))
            {
                string color = GetParameterValue<string>(parameters, "font_color", "");
                if (!string.IsNullOrEmpty(color))
                {
                    previewStyles.Add($"🎨 字体颜色: {color}");
                }
            }

            if (parameters.ContainsKey("background_color"))
            {
                string bgColor = GetParameterValue<string>(parameters, "background_color", "");
                if (!string.IsNullOrEmpty(bgColor))
                {
                    previewStyles.Add($"🖍️ 背景颜色: {(bgColor == "none" ? "无" : bgColor)}");
                }
            }

            if (parameters.ContainsKey("paragraph_spacing_before"))
            {
                int spacingBefore = GetParameterValue<int>(parameters, "paragraph_spacing_before", 0);
                previewStyles.Add($"📏 段前间距: {spacingBefore}磅");
            }

            if (parameters.ContainsKey("paragraph_spacing_after"))
            {
                int spacingAfter = GetParameterValue<int>(parameters, "paragraph_spacing_after", 0);
                previewStyles.Add($"📐 段后间距: {spacingAfter}磅");
            }

            if (parameters.ContainsKey("line_spacing"))
            {
                double lineSpacing = GetParameterValue<double>(parameters, "line_spacing", 1.0);
                previewStyles.Add($"📊 行间距: {lineSpacing}倍");
            }

            return previewStyles;
        }

        // 执行实际的样式修改
        private static string ExecuteStyleModification(Dictionary<string, object> parameters)
        {
            // 防止重复执行
            lock (_styleModificationLock)
            {
                if (_isStyleModificationInProgress)
                {
                    Debug.WriteLine("样式修改正在进行中，忽略重复请求");
                    return JsonConvert.SerializeObject(new
                    {
                        success = false,
                        message = "样式修改正在进行中，请等待完成后再试"
                    });
                }
                _isStyleModificationInProgress = true;
            }

            try
            {
                const int maxRetries = 3;
                const int retryDelayMs = 500;
                string result = null;

                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    try
                    {
                        Debug.WriteLine($"开始样式修改，尝试第 {attempt} 次...");

                        string textToFind = GetParameterValue<string>(parameters, "text_to_find", "");

                        // 检查文本长度 - Word COM Find对象有长度限制
                        if (string.IsNullOrEmpty(textToFind))
                        {
                            result = JsonConvert.SerializeObject(new
                            {
                                success = false,
                                message = "未提供要查找的文本"
                            });
                            break;
                        }

                        if (textToFind.Length > 255)
                        {
                            // Word Find对象通常限制在255个字符以内
                            result = JsonConvert.SerializeObject(new
                            {
                                success = false,
                                message = $"要查找的文本过长({textToFind.Length}个字符)，请缩短到255字符以内。建议使用关键词进行样式修改。"
                            });
                            break;
                        }

                        Debug.WriteLine($"样式修改目标文本长度: {textToFind.Length} 字符");
                        Debug.WriteLine($"样式修改目标文本: {textToFind.Substring(0, Math.Min(50, textToFind.Length))}...");

                        // 获取Word应用程序实例
                        dynamic wordApp = _markdownToWord.GetWordApplication();
                        if (wordApp == null)
                        {
                            result = JsonConvert.SerializeObject(new { success = false, message = "无法连接到Word应用程序" });
                            break;
                        }

                        Debug.WriteLine("成功连接到已打开的Word实例");

                        dynamic doc = _markdownToWord.GetActiveDocument(wordApp);
                        if (doc == null)
                        {
                            result = JsonConvert.SerializeObject(new { success = false, message = "无法获取活动文档" });
                            break;
                        }

                        // 等待Word空闲
                        System.Threading.Thread.Sleep(100);

                        // 查找文本 - 使用更灵活的匹配策略
                        dynamic range = doc.Content;
                        dynamic findObj = range.Find;
                        findObj.ClearFormatting(); // 清除之前的查找格式
                        findObj.Text = textToFind;
                        findObj.Forward = true;
                        findObj.Wrap = 1; // wdFindContinue

                        // 根据文本长度和内容调整匹配策略
                        if (textToFind.Length > 10 || textToFind.Contains("：") || textToFind.Contains(":"))
                        {
                            // 对于较长的文本或包含标点的文本，使用部分匹配
                            findObj.MatchWholeWord = false;
                        }
                        else
                        {
                            // 对于短文本，使用完整单词匹配
                            findObj.MatchWholeWord = true;
                        }
                        findObj.MatchCase = false;

                        int modifiedCount = 0;
                        List<string> appliedStyles = new List<string>();
                        const int maxMatches = 10; // 限制最大匹配数，避免过度匹配和重复

                        Debug.WriteLine($"开始查找文本: {textToFind}");

                        // 查找所有匹配项并应用样式
                        bool foundAny = false;
                        int lastFoundPosition = -1; // 记录上次找到的位置，避免重复

                        while (findObj.Execute() && modifiedCount < maxMatches)
                        {
                            foundAny = true;

                            // 检查是否和上次找到的位置相同，避免无限循环
                            int currentPosition = range.Start;
                            if (currentPosition == lastFoundPosition)
                            {
                                Debug.WriteLine($"检测到重复位置 {currentPosition}，停止查找以避免无限循环");
                                break;
                            }
                            lastFoundPosition = currentPosition;

                            try
                            {
                                ApplyTextStyles(range, parameters, appliedStyles);
                                modifiedCount++;
                                Debug.WriteLine($"成功修改第 {modifiedCount} 处文本，位置: {currentPosition}");

                                // 重新设置查找范围：从匹配文本末尾开始到文档末尾
                                int nextStartPosition = range.End + 1;
                                if (nextStartPosition >= doc.Content.End)
                                {
                                    Debug.WriteLine("已到达文档末尾，停止查找");
                                    break; // 已到文档末尾
                                }

                                // 重新创建查找范围，从下一个位置开始到文档末尾
                                range = doc.Range(nextStartPosition, doc.Content.End);
                                findObj = range.Find;
                                findObj.ClearFormatting();
                                findObj.Text = textToFind;
                                findObj.Forward = true;
                                findObj.Wrap = 0; // wdFindStop - 不要循环回到文档开头

                                // 应用相同的匹配策略
                                if (textToFind.Length > 10 || textToFind.Contains("：") || textToFind.Contains(":"))
                                {
                                    findObj.MatchWholeWord = false;
                                }
                                else
                                {
                                    findObj.MatchWholeWord = true;
                                }
                                findObj.MatchCase = false;

                                Debug.WriteLine($"重新设置查找范围：从位置 {nextStartPosition} 到文档末尾");
                            }
                            catch (System.Runtime.InteropServices.COMException comEx)
                            {
                                Debug.WriteLine($"应用样式时COM错误: {comEx.Message}");
                                if (comEx.HResult == unchecked((int)0x8001010A)) // RPC_E_SERVERCALL_RETRYLATER
                                {
                                    throw; // 重新抛出让外层重试
                                }
                                // 其他COM错误跳过这一个匹配项，重新设置查找范围
                                int skipPosition = range.End + 1;
                                if (skipPosition >= doc.Content.End)
                                {
                                    break;
                                }

                                range = doc.Range(skipPosition, doc.Content.End);
                                findObj = range.Find;
                                findObj.ClearFormatting();
                                findObj.Text = textToFind;
                                findObj.Forward = true;
                                findObj.Wrap = 0;

                                if (textToFind.Length > 10 || textToFind.Contains("：") || textToFind.Contains(":"))
                                {
                                    findObj.MatchWholeWord = false;
                                }
                                else
                                {
                                    findObj.MatchWholeWord = true;
                                }
                                findObj.MatchCase = false;
                                continue;
                            }
                        }

                        // 如果没有找到任何匹配项，尝试更宽松的查找
                        if (modifiedCount == 0 && !foundAny)
                        {
                            Debug.WriteLine("首次查找失败，尝试更宽松的匹配策略");

                            // 重置查找范围
                            range = doc.Content;
                            findObj = range.Find;
                            findObj.ClearFormatting();
                            findObj.Text = textToFind;
                            findObj.Forward = true;
                            findObj.Wrap = 1;
                            findObj.MatchWholeWord = false; // 使用更宽松的匹配
                            findObj.MatchCase = false;
                            findObj.MatchWildcards = false;

                            // 再次尝试查找
                            int fallbackLastPosition = -1;
                            while (findObj.Execute() && modifiedCount < maxMatches)
                            {
                                // 检查备用策略中的重复位置
                                int currentPos = range.Start;
                                if (currentPos == fallbackLastPosition)
                                {
                                    Debug.WriteLine($"备用策略检测到重复位置 {currentPos}，停止查找");
                                    break;
                                }
                                fallbackLastPosition = currentPos;

                                try
                                {
                                    ApplyTextStyles(range, parameters, appliedStyles);
                                    modifiedCount++;
                                    Debug.WriteLine($"备用策略成功修改第 {modifiedCount} 处文本，位置: {currentPos}");

                                    // 重新设置备用查找范围
                                    int fallbackNextPos = range.End + 1;
                                    if (fallbackNextPos >= doc.Content.End)
                                    {
                                        Debug.WriteLine("备用策略：已到达文档末尾");
                                        break;
                                    }

                                    range = doc.Range(fallbackNextPos, doc.Content.End);
                                    findObj = range.Find;
                                    findObj.ClearFormatting();
                                    findObj.Text = textToFind;
                                    findObj.Forward = true;
                                    findObj.Wrap = 0;
                                    findObj.MatchWholeWord = false;
                                    findObj.MatchCase = false;
                                    findObj.MatchWildcards = false;
                                }
                                catch (System.Runtime.InteropServices.COMException comEx)
                                {
                                    Debug.WriteLine($"备用策略应用样式时COM错误: {comEx.Message}");
                                    if (comEx.HResult == unchecked((int)0x8001010A))
                                    {
                                        throw;
                                    }
                                    // 备用策略COM错误处理：重新设置查找范围
                                    int fallbackSkipPos = range.End + 1;
                                    if (fallbackSkipPos >= doc.Content.End)
                                    {
                                        break;
                                    }

                                    range = doc.Range(fallbackSkipPos, doc.Content.End);
                                    findObj = range.Find;
                                    findObj.ClearFormatting();
                                    findObj.Text = textToFind;
                                    findObj.Forward = true;
                                    findObj.Wrap = 0;
                                    findObj.MatchWholeWord = false;
                                    findObj.MatchCase = false;
                                    findObj.MatchWildcards = false;
                                    continue;
                                }
                            }
                        }

                        if (modifiedCount == 0)
                        {
                            result = JsonConvert.SerializeObject(new { success = false, message = $"未找到文本: {textToFind}" });
                            break;
                        }

                        Debug.WriteLine($"样式修改完成，共修改 {modifiedCount} 处");
                        result = JsonConvert.SerializeObject(new
                        {
                            success = true,
                            message = $"成功修改了 {modifiedCount} 处文本的样式",
                            text_found = textToFind,
                            modified_count = modifiedCount,
                            applied_styles = appliedStyles
                        });
                        break;
                    }
                    catch (System.Runtime.InteropServices.COMException comEx) when (comEx.HResult == unchecked((int)0x8001010A))
                    {
                        Debug.WriteLine($"Word正在忙碌 (尝试 {attempt}/{maxRetries}): {comEx.Message}");

                        if (attempt < maxRetries)
                        {
                            // 等待后重试
                            System.Threading.Thread.Sleep(retryDelayMs * attempt);
                            continue;
                        }
                        else
                        {
                            result = JsonConvert.SerializeObject(new
                            {
                                success = false,
                                message = $"Word应用程序忙碌，请稍后再试。已重试 {maxRetries} 次"
                            });
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"执行样式修改时出错: {ex.Message}");
                        result = JsonConvert.SerializeObject(new { success = false, message = $"样式修改失败: {ex.Message}" });
                        break;
                    }
                }

                return result ?? JsonConvert.SerializeObject(new { success = false, message = "样式修改失败，未知错误" });
            }
            finally
            {
                // 确保锁被释放
                lock (_styleModificationLock)
                {
                    _isStyleModificationInProgress = false;
                    Debug.WriteLine("样式修改锁已释放");
                }
            }
        }

        // 应用文字样式
        private static void ApplyTextStyles(dynamic range, Dictionary<string, object> parameters, List<string> appliedStyles)
        {
            try
            {
                // 字体大小
                if (parameters.ContainsKey("font_size"))
                {
                    int fontSize = GetParameterValue<int>(parameters, "font_size", 0);
                    if (fontSize >= 8 && fontSize <= 72)
                    {
                        range.Font.Size = fontSize;
                        appliedStyles.Add($"字体大小: {fontSize}磅");
                    }
                }

                // 粗体
                if (parameters.ContainsKey("font_bold"))
                {
                    bool bold = GetParameterValue<bool>(parameters, "font_bold", false);
                    range.Font.Bold = bold;
                    appliedStyles.Add($"粗体: {(bold ? "是" : "否")}");
                }

                // 斜体
                if (parameters.ContainsKey("font_italic"))
                {
                    bool italic = GetParameterValue<bool>(parameters, "font_italic", false);
                    range.Font.Italic = italic;
                    appliedStyles.Add($"斜体: {(italic ? "是" : "否")}");
                }

                // 字体颜色
                if (parameters.ContainsKey("font_color"))
                {
                    string color = GetParameterValue<string>(parameters, "font_color", "").ToLower();
                    if (!string.IsNullOrEmpty(color))
                    {
                        int colorValue = GetWordColor(color);
                        if (colorValue != -1)
                        {
                            range.Font.Color = colorValue;
                            appliedStyles.Add($"字体颜色: {color}");
                        }
                    }
                }

                // 背景颜色（高亮）
                if (parameters.ContainsKey("background_color"))
                {
                    string bgColor = GetParameterValue<string>(parameters, "background_color", "").ToLower();
                    if (!string.IsNullOrEmpty(bgColor) && bgColor != "none")
                    {
                        int highlightValue = GetWordHighlightColor(bgColor);
                        if (highlightValue != -1)
                        {
                            range.HighlightColorIndex = highlightValue;
                            appliedStyles.Add($"背景色: {bgColor}");
                        }
                    }
                    else if (bgColor == "none")
                    {
                        range.HighlightColorIndex = 0; // wdNoHighlight
                        appliedStyles.Add("背景色: 无");
                    }
                }

                // 段落间距
                dynamic paragraph = range.Paragraphs[1];

                if (parameters.ContainsKey("paragraph_spacing_before"))
                {
                    int spacingBefore = GetParameterValue<int>(parameters, "paragraph_spacing_before", 0);
                    if (spacingBefore >= 0 && spacingBefore <= 100)
                    {
                        paragraph.SpaceBefore = spacingBefore;
                        appliedStyles.Add($"段前间距: {spacingBefore}磅");
                    }
                }

                if (parameters.ContainsKey("paragraph_spacing_after"))
                {
                    int spacingAfter = GetParameterValue<int>(parameters, "paragraph_spacing_after", 0);
                    if (spacingAfter >= 0 && spacingAfter <= 100)
                    {
                        paragraph.SpaceAfter = spacingAfter;
                        appliedStyles.Add($"段后间距: {spacingAfter}磅");
                    }
                }

                // 行间距
                if (parameters.ContainsKey("line_spacing"))
                {
                    double lineSpacing = GetParameterValue<double>(parameters, "line_spacing", 1.0);
                    if (lineSpacing >= 1.0 && lineSpacing <= 3.0)
                    {
                        paragraph.LineSpacing = lineSpacing * 12; // Word使用磅值
                        paragraph.LineSpacingRule = 5; // wdLineSpaceExactly
                        appliedStyles.Add($"行间距: {lineSpacing}倍");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"应用样式时出错: {ex.Message}");
                throw; // 重新抛出异常让上层处理
            }
        }

        // 获取Word颜色值
        private static int GetWordColor(string colorName)
        {
            switch (colorName.ToLower())
            {
                case "black": return 0x000000;
                case "red": return 0x0000FF;
                case "green": return 0x008000;
                case "blue": return 0xFF0000;
                case "yellow": return 0x00FFFF;
                case "orange": return 0x0080FF;
                case "purple": return 0x800080;
                case "gray": return 0x808080;
                case "white": return 0xFFFFFF;
                default: return -1;
            }
        }

        // 获取Word高亮颜色值
        private static int GetWordHighlightColor(string colorName)
        {
            switch (colorName.ToLower())
            {
                case "yellow": return 7; // wdYellow
                case "lightblue": return 11; // wdTurquoise
                case "lightgreen": return 4; // wdBrightGreen
                case "pink": return 5; // wdPink
                case "lightgray": return 16; // wdGray25
                case "white": return 8; // wdWhite
                default: return -1;
            }
        }

        // 判断是否为标题样式
        private static bool IsHeadingStyle(string styleName)
        {
            if (string.IsNullOrEmpty(styleName))
                return false;

            styleName = styleName.ToLower();

            // 检查中文标题样式
            if (styleName.Contains("标题") || styleName.Contains("heading"))
                return true;

            // 检查数字标题样式
            for (int i = 1; i <= 9; i++)
            {
                if (styleName.Contains($"标题 {i}") ||
                    styleName.Contains($"heading {i}") ||
                    styleName.Contains($"heading{i}") ||
                    styleName.Equals($"heading {i}") ||
                    styleName.Equals($"heading{i}"))
                {
                    return true;
                }
            }

            // 检查其他可能的标题样式名称
            string[] headingPatterns = {
                "heading", "标题", "title", "subtitle", "h1", "h2", "h3", "h4", "h5", "h6"
            };

            foreach (string pattern in headingPatterns)
            {
                if (styleName.Contains(pattern))
                    return true;
            }

            return false;
        }



        // 公开的直接执行方法，供UserControl1调用
        public static string ExecuteFormattedInsertDirectly(Dictionary<string, object> parameters)
        {
            try
            {
                Debug.WriteLine("=== ExecuteFormattedInsertDirectly 开始 ===");

                // 确保设置为实际执行模式
                var paramDict = new Dictionary<string, object>(parameters);
                paramDict["preview_only"] = false;

                Debug.WriteLine($"参数: preview_only = {paramDict["preview_only"]}");
                Debug.WriteLine($"参数: target_heading = {(paramDict.ContainsKey("target_heading") ? paramDict["target_heading"] : "未设置")}");
                Debug.WriteLine($"参数: content = {(paramDict.ContainsKey("content") ? paramDict["content"].ToString().Substring(0, Math.Min(50, paramDict["content"].ToString().Length)) : "未设置")}...");

                // 调用主要的FormattedInsertContent方法，它会处理preview_only参数
                string result = FormattedInsertContent(paramDict);

                Debug.WriteLine($"FormattedInsertContent 返回结果: {result.Substring(0, Math.Min(200, result.Length))}...");
                Debug.WriteLine("=== ExecuteFormattedInsertDirectly 结束 ===");

                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExecuteFormattedInsertDirectly 异常: {ex.Message}");
                return JsonConvert.SerializeObject(new { success = false, message = $"执行失败: {ex.Message}" });
            }
        }

        public static string ExecuteStyleModificationDirectly(Dictionary<string, object> parameters)
        {
            var paramDict = new Dictionary<string, object>(parameters);
            paramDict["preview_only"] = false;
            return ModifyTextStyle(paramDict);
        }







        // 注册文档统计工具
        private static void RegisterDocumentStatsTool()
        {
            OpenAIUtils.RegisterTool(
                "get_document_statistics",
                "获取当前Word文档的统计信息，包括字数、段落数、页数等",
                new Dictionary<string, object>(),
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => GetDocumentStatistics());
                });
        }



        // 注册获取选中文本工具
        private static void RegisterGetSelectedTextTool()
        {
            OpenAIUtils.RegisterTool(
                "get_selected_text",
                "获取当前在Word文档中选中的文本内容",
                new Dictionary<string, object>(),
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => GetSelectedText());
                });
        }





        // 注册获取图片工具
        private static void RegisterGetImagesTools()
        {
            OpenAIUtils.RegisterTool(
                "get_document_images",
                "获取文档中的所有图片信息，包括图片位置、大小、说明文字等",
                new Dictionary<string, object>
                {
                    { "include_details", new { type = "boolean", description = "是否包含详细信息如尺寸、格式等，默认为true" } }
                },
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => GetDocumentImages(parameters));
                });
        }

        // 注册获取公式工具
        private static void RegisterGetFormulasTools()
        {
            OpenAIUtils.RegisterTool(
                "get_document_formulas",
                "获取文档中数学公式的数量和位置信息。注意：此工具仅提供公式统计信息，不提取具体内容。",
                new Dictionary<string, object>
                {
                    { "formula_type", new { type = "string", description = "公式类型：all(所有), inline(行内), display(独立显示)，默认为all" } }
                },
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => GetDocumentFormulas(parameters));
                });
        }

        // 注册获取表格工具
        private static void RegisterGetTablesTools()
        {
            OpenAIUtils.RegisterTool(
                "get_document_tables",
                "获取文档中的所有表格信息，包括表格内容、行数、列数、位置等",
                new Dictionary<string, object>
                {
                    { "include_content", new { type = "boolean", description = "是否包含表格详细内容，默认为true" } },
                    { "table_index", new { type = "integer", description = "指定表格索引（从1开始），留空则获取所有表格" } }
                },
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => GetDocumentTables(parameters));
                });
        }

        // 注册获取文档标题列表工具（整体结构）
        private static void RegisterGetDocumentHeadingsTool()
        {
            OpenAIUtils.RegisterTool(
                "get_document_headings",
                "获取文档的标题列表（整体结构），返回所有标题的文本、级别、位置等信息。支持分页获取以提升性能。当用户询问文档结构、目录、大纲时应优先使用此工具。",
                new Dictionary<string, object>
                {
                    { "level", new { type = "integer", description = "可选，指定要获取的标题级别（1-9），不指定则返回所有级别的标题", minimum = 1, maximum = 9 } },
                    { "page", new { type = "integer", description = "可选，页码（从0开始）。不指定则返回所有标题。建议：超过20个标题时使用分页", minimum = 0, @default = -1 } },
                    { "page_size", new { type = "integer", description = "可选，每页数量（默认20）。仅在指定page时生效", minimum = 1, maximum = 100, @default = 20 } }
                },
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => GetDocumentHeadings(parameters));
                });
        }

        // 注册获取标题内容工具
        private static void RegisterGetHeadingContentTool()
        {
            OpenAIUtils.RegisterTool(
                "get_heading_content",
                "高效获取指定标题下的所有内容。使用智能匹配算法快速定位标题，无需遍历整个文档",
                new Dictionary<string, object>
                {
                    { "heading_text", new { type = "string", description = "要查找的标题文本（支持完全匹配、关键词匹配和部分匹配）" } },
                    { "match_level", new { type = "string", description = "匹配级别：exact(精确匹配)、keywords(关键词匹配)、fuzzy(模糊匹配)，默认为keywords", @enum = new[] { "exact", "keywords", "fuzzy" }, @default = "keywords" } },
                    { "max_content_length", new { type = "integer", description = "返回内容的最大长度（字符数），默认为2000，0表示无限制", @default = 2000, minimum = 0, maximum = 10000 } },
                    { "include_sub_headings", new { type = "boolean", description = "是否包含子标题下的内容，默认为true", @default = true } }
                },
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => GetHeadingContent(parameters));
                });
        }

        // 注册检查插入位置工具
        private static void RegisterCheckInsertPositionTool()
        {
            OpenAIUtils.RegisterTool(
                "check_insert_position",
                "检查指定标题的插入位置和上下文信息，获取标题下的现有内容以及前后相邻标题的内容作为参考",
                new Dictionary<string, object>
                {
                    { "target_heading", new { type = "string", description = "目标标题文本" } },
                    { "get_adjacent_context", new { type = "boolean", description = "是否获取前后相邻标题的内容作为上下文，默认为true", @default = true } },
                    { "max_context_length", new { type = "integer", description = "每个上下文内容的最大长度，默认为500字符", @default = 500, minimum = 100, maximum = 2000 } }
                },
                async (parameters) =>
                {
                    return await TaskAsync.Run(() => CheckInsertPosition(parameters));
                });
        }

        // 实现获取文档标题的功能（使用OutlineLevel属性，高效且支持取消）
        private static string GetDocumentHeadings(Dictionary<string, object> parameters)
        {
            try
            {
                var app = _markdownToWord.GetWordApplication();
                if (app == null)
                {
                    return "错误：无法连接到Word应用程序，请确保Word已打开";
                }

                var doc = _markdownToWord.GetActiveDocument(app);
                if (doc == null)
                {
                    return "错误：无法获取活动的Word文档";
                }

                int? targetLevel = null;
                if (parameters.ContainsKey("level") && parameters["level"] != null)
                {
                    if (int.TryParse(parameters["level"].ToString(), out int level))
                    {
                        targetLevel = level;
                    }
                }

                // 解析分页参数
                int page = -1;
                int pageSize = 20;
                if (parameters.ContainsKey("page") && parameters["page"] != null)
                {
                    if (int.TryParse(parameters["page"].ToString(), out int p))
                    {
                        page = p;
                    }
                }
                if (parameters.ContainsKey("page_size") && parameters["page_size"] != null)
                {
                    if (int.TryParse(parameters["page_size"].ToString(), out int ps))
                    {
                        pageSize = Math.Max(1, Math.Min(100, ps)); // 限制在1-100之间
                    }
                }

                var headings = new List<object>();
                var paragraphs = doc.Paragraphs;

                Debug.WriteLine($"开始使用OutlineLevel扫描 {paragraphs.Count} 个段落");

                int processedCount = 0;
                foreach (dynamic para in paragraphs)
                {
                    try
                    {
                        // 每处理100个段落检查一次是否需要取消
                        if (processedCount % 100 == 0)
                        {
                            // 检查是否有取消请求（通过检查全局状态或其他机制）
                            if (ShouldCancelOperation())
                            {
                                Debug.WriteLine("检测到取消请求，停止获取标题操作");
                                return JsonConvert.SerializeObject(new
                                {
                                    total_headings = headings.Count,
                                    target_level = targetLevel,
                                    headings = headings,
                                    cancelled = true,
                                    message = "操作已取消"
                                }, Formatting.Indented);
                            }
                        }
                        processedCount++;

                        // 直接检查OutlineLevel属性（高效且准确）
                        var outlineLevel = para.OutlineLevel;
                        if (outlineLevel != null && outlineLevel >= 1 && outlineLevel <= 9)
                        {
                            int level = (int)outlineLevel;

                            // 如果指定了级别，只返回该级别的标题
                            if (targetLevel.HasValue && level != targetLevel.Value)
                                continue;

                            string headingText = para.Range.Text?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(headingText))
                            {
                                headingText = headingText.Replace("\r", "").Replace("\n", "").Replace("\x07", "").Trim();

                                if (!string.IsNullOrEmpty(headingText))
                                {
                                    var pageNum = GetPageNumber(para.Range);

                                    // 获取样式名称（用于兼容性）
                                    string styleName = "";
                                    try
                                    {
                                        styleName = para.Style?.NameLocal ?? "";
                                    }
                                    catch (Exception styleEx)
                                    {
                                        Debug.WriteLine($"获取样式名称时出错: {styleEx.Message}");
                                    }

                                    headings.Add(new
                                    {
                                        text = headingText,
                                        level = level,
                                        position = para.Range.Start,
                                        pageNumber = pageNum,
                                        styleName = styleName,
                                        outlineLevel = level
                                    });

                                    Debug.WriteLine($"OutlineLevel标题: {headingText} (级别{level}, 页码{pageNum})");
                                }
                            }
                        }
                    }
                    catch (Exception paraEx)
                    {
                        Debug.WriteLine($"处理段落时出错: {paraEx.Message}");
                        continue; // 跳过有问题的段落，继续处理下一个
                    }
                }

                Debug.WriteLine($"标题扫描完成，共找到 {headings.Count} 个标题");

                if (headings.Count == 0)
                {
                    string levelInfo = targetLevel.HasValue ? $"级别 {targetLevel} 的" : "";
                    return $"未找到{levelInfo}标题。文档可能没有使用标题样式，或者没有{levelInfo}标题。";
                }

                // 实现分页逻辑
                object result;
                if (page >= 0)
                {
                    // 分页模式
                    int totalHeadings = headings.Count;
                    int totalPages = (int)Math.Ceiling((double)totalHeadings / pageSize);
                    int startIndex = page * pageSize;
                    
                    if (startIndex >= totalHeadings)
                    {
                        // 页码超出范围
                        result = new
                        {
                            total_headings = totalHeadings,
                            target_level = targetLevel,
                            current_page = page,
                            page_size = pageSize,
                            total_pages = totalPages,
                            has_more = false,
                            headings = new List<object>(),
                            message = $"页码超出范围。总共有 {totalPages} 页（从第0页开始）"
                        };
                    }
                    else
                    {
                        // 返回当前页的数据
                        var pagedHeadings = headings.Skip(startIndex).Take(pageSize).ToList();
                        bool hasMore = (startIndex + pageSize) < totalHeadings;
                        
                        result = new
                        {
                            total_headings = totalHeadings,
                            target_level = targetLevel,
                            current_page = page,
                            page_size = pageSize,
                            total_pages = totalPages,
                            has_more = hasMore,
                            headings = pagedHeadings,
                            message = $"已返回第 {page + 1}/{totalPages} 页（共 {totalHeadings} 个标题）"
                        };
                        
                        Debug.WriteLine($"分页返回: 页{page}, 每页{pageSize}, 返回{pagedHeadings.Count}个标题");
                    }
                }
                else
                {
                    // 非分页模式（向后兼容）
                    result = new
                    {
                        total_headings = headings.Count,
                        target_level = targetLevel,
                        headings = headings
                    };
                    
                    Debug.WriteLine($"全量返回: {headings.Count}个标题");
                }

                return JsonConvert.SerializeObject(result, Formatting.Indented);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档标题时出错: {ex.Message}");
                return $"获取文档标题失败: {ex.Message}";
            }
        }



        // 实现获取文档统计的功能
        private static string GetDocumentStatistics()
        {
            try
            {
                var app = _markdownToWord.GetWordApplication();
                if (app == null)
                {
                    return "错误：无法连接到Word应用程序，请确保Word已打开";
                }

                var doc = _markdownToWord.GetActiveDocument(app);
                if (doc == null)
                {
                    return "错误：无法获取活动的Word文档";
                }

                // 更新统计信息
                doc.Range().ComputeStatistics(WdStatistic.wdStatisticWords);

                var stats = new
                {
                    document_name = doc.Name,
                    word_count = doc.Range().ComputeStatistics(WdStatistic.wdStatisticWords),
                    character_count = doc.Range().ComputeStatistics(WdStatistic.wdStatisticCharacters),
                    character_count_no_spaces = doc.Range().ComputeStatistics(WdStatistic.wdStatisticCharactersWithSpaces),
                    paragraph_count = doc.Range().ComputeStatistics(WdStatistic.wdStatisticParagraphs),
                    line_count = doc.Range().ComputeStatistics(WdStatistic.wdStatisticLines),
                    page_count = doc.Range().ComputeStatistics(WdStatistic.wdStatisticPages),
                    comment_count = doc.Comments.Count,
                    table_count = doc.Tables.Count,
                    section_count = doc.Sections.Count
                };

                return JsonConvert.SerializeObject(stats, Formatting.Indented);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档统计时出错: {ex.Message}");
                return $"获取文档统计失败: {ex.Message}";
            }
        }



        // 实现获取选中文本功能
        private static string GetSelectedText()
        {
            try
            {
                var app = _markdownToWord.GetWordApplication();
                if (app == null)
                {
                    return "错误：无法连接到Word应用程序，请确保Word已打开";
                }

                var selection = app.Selection;
                if (selection.Text == null || string.IsNullOrWhiteSpace(selection.Text))
                {
                    return "当前没有选中任何文本。请先在Word文档中选择要操作的文本。";
                }

                var selectedInfo = new
                {
                    text = selection.Text.Trim(),
                    start_position = selection.Start,
                    end_position = selection.End,
                    length = selection.Text.Length,
                    word_count = selection.Range.ComputeStatistics(WdStatistic.wdStatisticWords),
                    paragraph_count = selection.Range.ComputeStatistics(WdStatistic.wdStatisticParagraphs)
                };

                return JsonConvert.SerializeObject(selectedInfo, Formatting.Indented);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取选中文本时出错: {ex.Message}");
                return $"获取选中文本失败: {ex.Message}";
            }
        }





        // 实现获取文档图片功能
        private static string GetDocumentImages(Dictionary<string, object> parameters)
        {
            try
            {
                var app = _markdownToWord.GetWordApplication();
                if (app == null)
                {
                    return "错误：无法连接到Word应用程序，请确保Word已打开";
                }

                var doc = _markdownToWord.GetActiveDocument(app);
                if (doc == null)
                {
                    return "错误：无法获取活动的Word文档";
                }

                bool includeDetails = parameters.ContainsKey("include_details") ?
                    bool.Parse(parameters["include_details"].ToString()) : true;

                var images = new List<object>();

                Debug.WriteLine($"开始扫描文档图片，包含详细信息: {includeDetails}");

                // 扫描内联图片（InlineShapes）
                for (int i = 1; i <= doc.InlineShapes.Count; i++)
                {
                    try
                    {
                        var shape = doc.InlineShapes[i];

                        // 安全的类型比较
                        bool isPicture = false;
                        try
                        {
                            var shapeType = (WdInlineShapeType)shape.Type;
                            isPicture = shapeType == WdInlineShapeType.wdInlineShapePicture;
                        }
                        catch (Exception typeEx)
                        {
                            Debug.WriteLine($"检查内联图片类型时出错: {typeEx.Message}");
                            // 尝试其他方法判断是否为图片
                            try
                            {
                                // 如果有Width和Height属性，可能是图片
                                var testWidth = shape.Width;
                                isPicture = true;
                            }
                            catch
                            {
                                isPicture = false;
                            }
                        }

                        if (isPicture)
                        {
                            var altText = "";
                            var width = 0.0;
                            var height = 0.0;

                            if (includeDetails)
                            {
                                try
                                {
                                    width = Math.Round(shape.Width, 2);
                                    height = Math.Round(shape.Height, 2);
                                    altText = shape.AlternativeText ?? "";
                                }
                                catch (Exception detailEx)
                                {
                                    Debug.WriteLine($"获取内联图片详细信息时出错: {detailEx.Message}");
                                }
                            }

                            var imageInfo = new
                            {
                                index = i,
                                type = "inline_picture",
                                position = shape.Range.Start,
                                width = width,
                                height = height,
                                alt_text = altText,
                                page_number = GetPageNumber(shape.Range)
                            };
                            images.Add(imageInfo);
                            Debug.WriteLine($"发现内联图片 {i}: 位置{shape.Range.Start}");
                        }
                    }
                    catch (Exception shapeEx)
                    {
                        Debug.WriteLine($"处理内联图片 {i} 时出错: {shapeEx.Message}");
                    }
                }

                // 扫描浮动图片（Shapes）
                for (int i = 1; i <= doc.Shapes.Count; i++)
                {
                    try
                    {
                        var shape = doc.Shapes[i];

                        // 安全的类型比较
                        bool isPicture = false;
                        try
                        {
                            var shapeType = (Microsoft.Office.Core.MsoShapeType)shape.Type;
                            isPicture = shapeType == Microsoft.Office.Core.MsoShapeType.msoPicture;
                        }
                        catch (Exception typeEx)
                        {
                            Debug.WriteLine($"检查浮动图片类型时出错: {typeEx.Message}");
                            // 尝试其他方法判断是否为图片
                            try
                            {
                                // 如果有Width和Height属性，且名称包含图片相关信息，可能是图片
                                var testWidth = shape.Width;
                                var shapeName = shape.Name ?? "";
                                isPicture = shapeName.Contains("图片") || shapeName.Contains("Picture") || shapeName.Contains("Image");
                            }
                            catch
                            {
                                isPicture = false;
                            }
                        }

                        if (isPicture)
                        {
                            var shapeName = $"图片{i}";
                            var altText = "";
                            var width = 0.0;
                            var height = 0.0;
                            var left = 0.0;
                            var top = 0.0;

                            try
                            {
                                shapeName = shape.Name ?? $"图片{i}";
                            }
                            catch { }

                            if (includeDetails)
                            {
                                try
                                {
                                    width = Math.Round(shape.Width, 2);
                                    height = Math.Round(shape.Height, 2);
                                    left = Math.Round(shape.Left, 2);
                                    top = Math.Round(shape.Top, 2);
                                    altText = shape.AlternativeText ?? "";
                                }
                                catch (Exception detailEx)
                                {
                                    Debug.WriteLine($"获取浮动图片详细信息时出错: {detailEx.Message}");
                                }
                            }

                            var imageInfo = new
                            {
                                index = i,
                                type = "floating_picture",
                                name = shapeName,
                                width = width,
                                height = height,
                                left = left,
                                top = top,
                                alt_text = altText
                            };
                            images.Add(imageInfo);
                            Debug.WriteLine($"发现浮动图片 {i}: {shapeName}");
                        }
                    }
                    catch (Exception shapeEx)
                    {
                        Debug.WriteLine($"处理浮动图片 {i} 时出错: {shapeEx.Message}");
                    }
                }

                var result = new
                {
                    total_images = images.Count,
                    inline_images = doc.InlineShapes.Count,
                    floating_images = doc.Shapes.Count,
                    include_details = includeDetails,
                    images = images
                };

                Debug.WriteLine($"图片扫描完成，共找到 {images.Count} 个图片");
                return JsonConvert.SerializeObject(result, Formatting.Indented);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档图片时出错: {ex.Message}");
                return $"获取文档图片失败: {ex.Message}";
            }
        }

        // 实现获取文档公式功能
        private static string GetDocumentFormulas(Dictionary<string, object> parameters)
        {
            try
            {
                var app = _markdownToWord.GetWordApplication();
                if (app == null)
                {
                    return "错误：无法连接到Word应用程序，请确保Word已打开";
                }

                var doc = _markdownToWord.GetActiveDocument(app);
                if (doc == null)
                {
                    return "错误：无法获取活动的Word文档";
                }

                string formulaType = parameters.ContainsKey("formula_type") ?
                    parameters["formula_type"].ToString().ToLower() : "all";

                var formulas = new List<object>();

                Debug.WriteLine($"开始扫描文档公式，类型: {formulaType}");

                // 扫描OMaths集合（Office Math）
                var oMaths = doc.Range().OMaths;
                for (int i = 1; i <= oMaths.Count; i++)
                {
                    try
                    {
                        var oMath = oMaths[i];

                        // 修复COM类型比较问题
                        string mathType = "display"; // 默认为独立公式
                        try
                        {
                            // 安全的类型比较
                            var omathType = (WdOMathType)oMath.Type;
                            mathType = omathType == WdOMathType.wdOMathInline ? "inline" : "display";
                        }
                        catch (Exception typeEx)
                        {
                            Debug.WriteLine($"获取公式类型时出错: {typeEx.Message}，使用默认类型");
                        }

                        // 根据过滤条件决定是否包含
                        if (formulaType == "all" || formulaType == mathType)
                        {
                            var formulaInfo = new
                            {
                                index = i,
                                type = mathType,
                                position = oMath.Range.Start,
                                page_number = GetPageNumber(oMath.Range),
                                note = "具体公式内容请在Word中查看"
                            };
                            formulas.Add(formulaInfo);

                            Debug.WriteLine($"发现{mathType}公式 {i}: 位置{oMath.Range.Start}");
                        }
                    }
                    catch (Exception mathEx)
                    {
                        Debug.WriteLine($"处理公式 {i} 时出错: {mathEx.Message}");
                        Debug.WriteLine($"错误堆栈: {mathEx.StackTrace}");
                    }
                }

                var result = new
                {
                    total_formulas = formulas.Count,
                    formula_type_filter = formulaType,
                    total_omath_objects = oMaths.Count,
                    formulas = formulas,
                    disclaimer = "此工具仅提供公式位置和数量统计。由于Word COM兼容性限制，不提供公式内容提取。请直接在Word文档中查看具体公式内容。",
                    suggestion = "要编辑或复制公式，请在Word中选中公式后使用右键菜单或公式工具。"
                };

                Debug.WriteLine($"公式扫描完成，共找到 {formulas.Count} 个公式");
                return JsonConvert.SerializeObject(result, Formatting.Indented);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档公式时出错: {ex.Message}");
                return $"获取文档公式失败: {ex.Message}";
            }
        }

        // 实现获取文档表格功能
        private static string GetDocumentTables(Dictionary<string, object> parameters)
        {
            try
            {
                var app = _markdownToWord.GetWordApplication();
                if (app == null)
                {
                    return "错误：无法连接到Word应用程序，请确保Word已打开";
                }

                var doc = _markdownToWord.GetActiveDocument(app);
                if (doc == null)
                {
                    return "错误：无法获取活动的Word文档";
                }

                bool includeContent = parameters.ContainsKey("include_content") ?
                    bool.Parse(parameters["include_content"].ToString()) : true;

                int? targetIndex = null;
                if (parameters.ContainsKey("table_index") && parameters["table_index"] != null)
                {
                    if (int.TryParse(parameters["table_index"].ToString(), out int index))
                    {
                        targetIndex = index;
                    }
                }

                var tables = new List<object>();

                Debug.WriteLine($"开始扫描文档表格，包含内容: {includeContent}, 目标索引: {targetIndex}");

                for (int i = 1; i <= doc.Tables.Count; i++)
                {
                    // 如果指定了表格索引，只处理该表格
                    if (targetIndex.HasValue && i != targetIndex.Value)
                        continue;

                    try
                    {
                        var table = doc.Tables[i];
                        var tableData = new
                        {
                            index = i,
                            rows = table.Rows.Count,
                            columns = table.Columns.Count,
                            position = table.Range.Start,
                            title = GetTableTitle(table),
                            page_number = GetPageNumber(table.Range),
                            content = includeContent ? GetTableContent(table) : null
                        };

                        tables.Add(tableData);
                        Debug.WriteLine($"处理表格 {i}: {table.Rows.Count}行 x {table.Columns.Count}列");
                    }
                    catch (Exception tableEx)
                    {
                        Debug.WriteLine($"处理表格 {i} 时出错: {tableEx.Message}");
                    }
                }

                var result = new
                {
                    total_tables = tables.Count,
                    document_total_tables = doc.Tables.Count,
                    include_content = includeContent,
                    target_index = targetIndex,
                    tables = tables
                };

                Debug.WriteLine($"表格扫描完成，共处理 {tables.Count} 个表格");
                return JsonConvert.SerializeObject(result, Formatting.Indented);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档表格时出错: {ex.Message}");
                return $"获取文档表格失败: {ex.Message}";
            }
        }

        // 辅助方法：获取页码
        private static int GetPageNumber(Range range)
        {
            try
            {
                return range.Information[WdInformation.wdActiveEndPageNumber];
            }
            catch
            {
                return 1; // 默认返回第1页
            }
        }

        // 辅助方法：获取表格标题
        private static string GetTableTitle(Table table)
        {
            try
            {
                // 尝试获取表格前面的段落作为标题
                var rangeBefore = table.Range;
                rangeBefore.MoveStart(WdUnits.wdParagraph, -1);
                var titleText = rangeBefore.Paragraphs[1].Range.Text.Trim();

                if (titleText.Length > 50)
                    titleText = titleText.Substring(0, 50) + "...";

                return titleText;
            }
            catch
            {
                return $"表格{table.Range.Start}";
            }
        }

        // 辅助方法：获取表格内容
        private static object GetTableContent(Table table)
        {
            try
            {
                var rows = new List<List<string>>();

                for (int r = 1; r <= Math.Min(table.Rows.Count, 10); r++) // 最多获取10行避免太大
                {
                    var row = new List<string>();
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        try
                        {
                            var cellText = table.Cell(r, c).Range.Text.Replace("\r", "").Replace("\x07", "").Trim();
                            if (cellText.Length > 100)
                                cellText = cellText.Substring(0, 100) + "...";
                            row.Add(cellText);
                        }
                        catch
                        {
                            row.Add("[无法读取]");
                        }
                    }
                    rows.Add(row);
                }

                return new
                {
                    headers = rows.Count > 0 ? rows[0] : new List<string>(),
                    data_rows = rows.Skip(1).ToList(),
                    note = table.Rows.Count > 10 ? $"表格共{table.Rows.Count}行，仅显示前10行" : ""
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取表格内容时出错: {ex.Message}");
                return new { error = $"获取表格内容失败: {ex.Message}" };
            }
        }

        // 实现检查插入位置功能
        private static string CheckInsertPosition(Dictionary<string, object> parameters)
        {
            try
            {
                Debug.WriteLine("开始检查插入位置和上下文...");

                // 获取参数
                string targetHeading = GetParameterValue<string>(parameters, "target_heading", "");
                bool getAdjacentContext = GetParameterValue<bool>(parameters, "get_adjacent_context", true);
                int maxContextLength = GetParameterValue<int>(parameters, "max_context_length", 500);

                if (string.IsNullOrEmpty(targetHeading))
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "目标标题不能为空" });
                }

                var app = _markdownToWord.GetWordApplication();
                if (app == null)
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "无法连接到Word应用程序" });
                }

                var doc = _markdownToWord.GetActiveDocument(app);
                if (doc == null)
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "无法获取活动文档" });
                }

                // 查找目标标题
                var headingResult = FindHeadingEfficiently(doc, targetHeading, "keywords");
                if (!headingResult.found)
                {
                    return JsonConvert.SerializeObject(new
                    {
                        success = false,
                        message = $"未找到标题: {targetHeading}",
                        suggestions = headingResult.suggestions ?? new string[0]
                    });
                }

                var targetHeadingInfo = headingResult.headingInfo;
                Debug.WriteLine($"找到目标标题: {targetHeadingInfo.text} (级别: {targetHeadingInfo.level})");

                // 获取目标标题下的现有内容
                string existingContent = ExtractContentUnderHeading(doc, targetHeadingInfo, true, maxContextLength);
                bool hasExistingContent = !string.IsNullOrWhiteSpace(existingContent);

                var result = new
                {
                    success = true,
                    target_heading = new
                    {
                        text = targetHeadingInfo.text,
                        level = targetHeadingInfo.level,
                        position = targetHeadingInfo.position,
                        has_existing_content = hasExistingContent,
                        existing_content = hasExistingContent ? existingContent : ""
                    },
                    adjacent_context = getAdjacentContext ? GetAdjacentHeadingsContext(doc, targetHeadingInfo, maxContextLength) : null,
                    insert_recommendation = hasExistingContent ? "append" : "create_new"
                };

                return JsonConvert.SerializeObject(result, Formatting.Indented);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"检查插入位置时出错: {ex.Message}");
                return JsonConvert.SerializeObject(new { success = false, message = $"检查插入位置失败: {ex.Message}" });
            }
        }

        // 获取相邻标题的上下文
        private static object GetAdjacentHeadingsContext(dynamic doc, dynamic targetHeadingInfo, int maxLength)
        {
            try
            {
                var allHeadings = new List<dynamic>();
                int targetPosition = targetHeadingInfo.position;
                int targetLevel = targetHeadingInfo.level;

                // 收集所有标题
                foreach (dynamic para in doc.Paragraphs)
                {
                    try
                    {
                        var outlineLevel = para.OutlineLevel;
                        if (outlineLevel != null && outlineLevel >= 1 && outlineLevel <= 9)
                        {
                            string text = para.Range.Text?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(text))
                            {
                                text = text.Replace("\r", "").Replace("\n", "").Trim();
                                if (!string.IsNullOrEmpty(text))
                                {
                                    allHeadings.Add(new
                                    {
                                        text = text,
                                        level = (int)outlineLevel,
                                        position = para.Range.Start,
                                        range = para.Range
                                    });
                                }
                            }
                        }
                    }
                    catch { continue; }
                }

                // 按位置排序
                allHeadings.Sort((a, b) => a.position.CompareTo(b.position));

                // 找到目标标题在列表中的索引
                int targetIndex = -1;
                for (int i = 0; i < allHeadings.Count; i++)
                {
                    if (allHeadings[i].position == targetPosition)
                    {
                        targetIndex = i;
                        break;
                    }
                }

                if (targetIndex == -1)
                {
                    return new { previous = (object)null, next = (object)null };
                }

                // 获取前一个同级或更高级标题
                dynamic previousHeading = null;
                for (int i = targetIndex - 1; i >= 0; i--)
                {
                    if (allHeadings[i].level <= targetLevel)
                    {
                        previousHeading = allHeadings[i];
                        break;
                    }
                }

                // 获取后一个同级或更高级标题
                dynamic nextHeading = null;
                for (int i = targetIndex + 1; i < allHeadings.Count; i++)
                {
                    if (allHeadings[i].level <= targetLevel)
                    {
                        nextHeading = allHeadings[i];
                        break;
                    }
                }

                // 提取上下文内容
                object previousContext = null;
                object nextContext = null;

                if (previousHeading != null)
                {
                    string prevContent = ExtractContentUnderHeading(doc, previousHeading, true, maxLength);
                    previousContext = new
                    {
                        heading = previousHeading.text,
                        level = previousHeading.level,
                        content = prevContent,
                        has_content = !string.IsNullOrWhiteSpace(prevContent)
                    };
                }

                if (nextHeading != null)
                {
                    string nextContent = ExtractContentUnderHeading(doc, nextHeading, true, maxLength);
                    nextContext = new
                    {
                        heading = nextHeading.text,
                        level = nextHeading.level,
                        content = nextContent,
                        has_content = !string.IsNullOrWhiteSpace(nextContent)
                    };
                }

                return new { previous = previousContext, next = nextContext };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取相邻标题上下文时出错: {ex.Message}");
                return new { previous = (object)null, next = (object)null, error = ex.Message };
            }
        }

        // 实现获取标题内容功能 - 高效且智能匹配
        private static string GetHeadingContent(Dictionary<string, object> parameters)
        {
            try
            {
                Debug.WriteLine("开始执行高效标题内容获取...");

                // 获取参数
                string headingText = GetParameterValue<string>(parameters, "heading_text", "");
                string matchLevel = GetParameterValue<string>(parameters, "match_level", "keywords");
                int maxLength = GetParameterValue<int>(parameters, "max_content_length", 2000);
                bool includeSubHeadings = GetParameterValue<bool>(parameters, "include_sub_headings", true);

                // 空参回退：如果未提供标题文本，返回文档整体结构（标题列表）
                if (string.IsNullOrEmpty(headingText))
                {
                    Debug.WriteLine("未提供标题文本，回退为返回文档整体结构");
                    NotifyProgress("未提供标题文本，返回文档标题列表");
                    
                    // 调用 GetDocumentHeadings 获取文档结构
                    string structureResult = GetDocumentHeadings(new Dictionary<string, object>());
                    
                    // 尝试解析结果并添加 mode 标记
                    try
                    {
                        var parsedResult = JsonConvert.DeserializeObject<dynamic>(structureResult);
                        if (parsedResult != null && parsedResult.total_headings != null)
                        {
                            var enhancedResult = new
                            {
                                success = true,
                                mode = "document_structure",
                                message = "未提供标题文本，已返回文档整体结构（标题列表）",
                                total = (int)parsedResult.total_headings,
                                headings = parsedResult.headings
                            };
                            return JsonConvert.SerializeObject(enhancedResult, Formatting.Indented);
                        }
                    }
                    catch (Exception parseEx)
                    {
                        Debug.WriteLine($"解析文档结构时出错: {parseEx.Message}");
                    }
                    
                    // 如果解析失败，返回原始结果包装
                    return JsonConvert.SerializeObject(new 
                    { 
                        success = true,
                        mode = "document_structure",
                        message = "未提供标题文本，已返回文档整体结构（标题列表）",
                        raw_result = structureResult
                    });
                }

                Debug.WriteLine($"查找标题: {headingText}, 匹配级别: {matchLevel}, 最大长度: {maxLength}");
                NotifyProgress($"查找标题: {headingText}, 匹配级别: {matchLevel}, 最大长度: {maxLength}");

                var app = _markdownToWord.GetWordApplication();
                if (app == null)
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "无法连接到Word应用程序" });
                }

                NotifyProgress("成功连接到已打开的Word实例");

                var doc = _markdownToWord.GetActiveDocument(app);
                if (doc == null)
                {
                    return JsonConvert.SerializeObject(new { success = false, message = "无法获取活动文档" });
                }

                // 高效查找目标标题
                var headingResult = FindHeadingEfficiently(doc, headingText, matchLevel);
                if (!headingResult.found)
                {
                    return JsonConvert.SerializeObject(new
                    {
                        success = false,
                        message = $"未找到标题: {headingText}",
                        suggestions = headingResult.suggestions ?? new string[0]
                    });
                }

                Debug.WriteLine($"找到标题: {headingResult.headingInfo.text} (级别: {headingResult.headingInfo.level})");
                NotifyProgress($"找到标题: {headingResult.headingInfo.text} (级别: {headingResult.headingInfo.level})");

                // 获取标题下的内容
                NotifyProgress($"开始提取标题下的内容: {headingResult.headingInfo.text}");
                string content = ExtractContentUnderHeading(doc, headingResult.headingInfo, includeSubHeadings, maxLength);
                NotifyProgress($"提取内容完成，长度: {content.Length} 字符");

                var result = new
                {
                    success = true,
                    found_heading = headingResult.headingInfo,
                    content = content,
                    content_length = content.Length,
                    truncated = maxLength > 0 && content.Length >= maxLength,
                    match_level = matchLevel,
                    include_sub_headings = includeSubHeadings
                };

                return JsonConvert.SerializeObject(result, Formatting.Indented);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取标题内容时出错: {ex.Message}");
                return JsonConvert.SerializeObject(new { success = false, message = $"获取标题内容失败: {ex.Message}" });
            }
        }

        // 高效查找标题的结果结构
        private class HeadingFindResult
        {
            public bool found;
            public dynamic headingInfo;
            public string[] suggestions;
        }

        // 高效查找标题 - 优化版本，减少全文档遍历
        private static HeadingFindResult FindHeadingEfficiently(dynamic doc, string targetHeading, string matchLevel)
        {
            var result = new HeadingFindResult { found = false, headingInfo = null, suggestions = new string[0] };
            var candidateHeadings = new List<dynamic>();
            var allHeadings = new List<string>();

            try
            {
                Debug.WriteLine($"开始高效查找标题: {targetHeading} (匹配级别: {matchLevel})");
                NotifyProgress($"开始高效查找标题: {targetHeading} (匹配级别: {matchLevel})");

                // 预处理目标标题
                string normalizedTarget = NormalizeText(targetHeading);
                string[] keywords = ExtractKeywords(targetHeading);

                Debug.WriteLine($"提取关键词: [{string.Join(", ", keywords)}]");
                NotifyProgress($"提取关键词: [{string.Join(", ", keywords)}]");

                // 使用OutlineLevel快速遍历标题（只遍历标题，不遍历所有段落）
                int headingCount = 0;
                int checkedCount = 0;
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                foreach (dynamic para in doc.Paragraphs)
                {
                    try
                    {
                        checkedCount++;

                        // 每检查500个段落报告一次进度
                        if (checkedCount % 500 == 0)
                        {
                            Debug.WriteLine($"已检查 {checkedCount} 个段落，找到 {headingCount} 个标题, 耗时: {stopwatch.ElapsedMilliseconds}ms");
                        }

                        // 性能保护：超过10秒或检查超过5000个段落就停止
                        if (stopwatch.ElapsedMilliseconds > 10000 || checkedCount > 5000)
                        {
                            Debug.WriteLine($"搜索超时或达到段落限制，停止搜索");
                            break;
                        }

                        // 检查是否为标题
                        var outlineLevel = para.OutlineLevel;
                        if (outlineLevel == null || outlineLevel < 1 || outlineLevel > 9)
                            continue;

                        headingCount++;
                        string paraText = para.Range.Text?.Trim() ?? "";
                        if (string.IsNullOrEmpty(paraText))
                            continue;

                        // 清理标题文本
                        paraText = paraText.Replace("\r", "").Replace("\n", "").Replace("\x07", "").Trim();
                        if (string.IsNullOrEmpty(paraText))
                            continue;

                        allHeadings.Add(paraText);
                        string normalizedPara = NormalizeText(paraText);

                        // 根据匹配级别进行匹配
                        bool isMatch = false;
                        double matchScore = 0.0;

                        switch (matchLevel.ToLower())
                        {
                            case "exact":
                                isMatch = normalizedPara == normalizedTarget ||
                                         paraText.Equals(targetHeading, StringComparison.OrdinalIgnoreCase);
                                matchScore = isMatch ? 1.0 : 0.0;
                                break;

                            case "keywords":
                                matchScore = CalculateKeywordMatchScore(paraText, keywords);
                                isMatch = matchScore >= 0.6; // 60%关键词匹配度
                                break;

                            case "fuzzy":
                                matchScore = CalculateFuzzyMatchScore(normalizedPara, normalizedTarget);
                                isMatch = matchScore >= 0.4; // 40%模糊匹配度
                                break;

                            default:
                                // 默认关键词匹配
                                matchScore = CalculateKeywordMatchScore(paraText, keywords);
                                isMatch = matchScore >= 0.6;
                                break;
                        }

                        if (isMatch)
                        {
                            var headingInfo = new
                            {
                                text = paraText,
                                level = (int)outlineLevel,
                                position = para.Range.Start,
                                matchScore = matchScore,
                                range = para.Range
                            };

                            candidateHeadings.Add(headingInfo);
                            Debug.WriteLine($"找到候选标题: {paraText} (匹配度: {matchScore:F2})");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"处理段落时出错: {ex.Message}");
                        continue;
                    }
                }

                stopwatch.Stop();
                Debug.WriteLine($"标题搜索完成 - 检查了{checkedCount}个段落，找到{headingCount}个标题，{candidateHeadings.Count}个候选标题，耗时: {stopwatch.ElapsedMilliseconds}ms");
                NotifyProgress($"标题搜索完成 - 检查了{checkedCount}个段落，找到{headingCount}个标题，{candidateHeadings.Count}个候选标题，耗时: {stopwatch.ElapsedMilliseconds}ms");

                if (candidateHeadings.Count > 0)
                {
                    // 选择最佳匹配（按匹配度排序）
                    var bestMatch = candidateHeadings
                        .OrderByDescending(h => h.matchScore)
                        .ThenBy(h => h.position) // 位置靠前优先
                        .First();

                    result.found = true;
                    result.headingInfo = bestMatch;
                    Debug.WriteLine($"选择最佳匹配: {bestMatch.text} (匹配度: {bestMatch.matchScore:F2})");
                    NotifyProgress($"选择最佳匹配: {bestMatch.text} (匹配度: {bestMatch.matchScore:F2})");
                }
                else
                {
                    // 生成建议
                    result.suggestions = GenerateHeadingSuggestions(allHeadings, targetHeading, 5);
                    Debug.WriteLine($"未找到匹配标题，生成了 {result.suggestions.Length} 个建议");
                }

                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"高效查找标题时出错: {ex.Message}");
                result.suggestions = new string[] { "搜索过程中发生错误，请检查标题文本或尝试其他关键词" };
                return result;
            }
        }

        // 文本标准化
        private static string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            return text.ToLower()
                      .Replace(" ", "")
                      .Replace("　", "") // 全角空格
                      .Replace("\t", "")
                      .Replace(":", "")
                      .Replace("：", "")
                      .Replace("-", "")
                      .Replace("–", "")
                      .Replace("—", "")
                      .Trim();
        }

        // 提取关键词
        private static string[] ExtractKeywords(string text)
        {
            if (string.IsNullOrEmpty(text)) return new string[0];

            var keywords = text.Split(new char[] { ' ', '　', '\t', ':', '：', '-', '–', '—', '(', ')', '（', '）', '[', ']', '【', '】' },
                                     StringSplitOptions.RemoveEmptyEntries)
                              .Where(k => k.Length > 1) // 过滤单字符
                              .Select(k => k.Trim())
                              .Where(k => !string.IsNullOrEmpty(k))
                              .Distinct()
                              .ToArray();

            return keywords;
        }

        // 计算关键词匹配分数
        private static double CalculateKeywordMatchScore(string text, string[] keywords)
        {
            if (keywords.Length == 0) return 0.0;

            string normalizedText = NormalizeText(text);
            int matchedCount = 0;

            foreach (string keyword in keywords)
            {
                string normalizedKeyword = NormalizeText(keyword);
                if (normalizedText.Contains(normalizedKeyword))
                {
                    matchedCount++;
                }
            }

            return (double)matchedCount / keywords.Length;
        }

        // 计算模糊匹配分数（简化的Levenshtein距离）
        private static double CalculateFuzzyMatchScore(string text1, string text2)
        {
            if (string.IsNullOrEmpty(text1) || string.IsNullOrEmpty(text2))
                return 0.0;

            int maxLen = Math.Max(text1.Length, text2.Length);
            if (maxLen == 0) return 1.0;

            int distance = LevenshteinDistance(text1, text2);
            return 1.0 - (double)distance / maxLen;
        }

        // Levenshtein距离算法（简化版）
        private static int LevenshteinDistance(string source, string target)
        {
            if (string.IsNullOrEmpty(source)) return target?.Length ?? 0;
            if (string.IsNullOrEmpty(target)) return source.Length;

            int[,] matrix = new int[source.Length + 1, target.Length + 1];

            for (int i = 0; i <= source.Length; i++)
                matrix[i, 0] = i;
            for (int j = 0; j <= target.Length; j++)
                matrix[0, j] = j;

            for (int i = 1; i <= source.Length; i++)
            {
                for (int j = 1; j <= target.Length; j++)
                {
                    int cost = source[i - 1] == target[j - 1] ? 0 : 1;
                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost);
                }
            }

            return matrix[source.Length, target.Length];
        }

        // 生成标题建议
        private static string[] GenerateHeadingSuggestions(List<string> allHeadings, string targetHeading, int maxSuggestions)
        {
            if (allHeadings.Count == 0) return new string[0];

            string[] keywords = ExtractKeywords(targetHeading);

            var suggestions = allHeadings
                .Select(heading => new {
                    text = heading,
                    score = CalculateKeywordMatchScore(heading, keywords) +
                           CalculateFuzzyMatchScore(NormalizeText(heading), NormalizeText(targetHeading)) * 0.3
                })
                .Where(s => s.score > 0.1) // 至少有一些相似度
                .OrderByDescending(s => s.score)
                .Take(maxSuggestions)
                .Select(s => s.text)
                .ToArray();

            return suggestions;
        }

        // 提取标题下的内容
        private static string ExtractContentUnderHeading(dynamic doc, dynamic headingInfo, bool includeSubHeadings, int maxLength)
        {
            try
            {
                Debug.WriteLine($"开始提取标题下的内容: {headingInfo.text}");

                var content = new StringBuilder();
                dynamic headingRange = headingInfo.range;
                int headingLevel = headingInfo.level;
                int headingEnd = headingRange.End;

                // 从标题后开始读取内容
                int contentStart = headingEnd;
                int contentEnd = doc.Content.End;

                // 查找下一个同级或更高级的标题，确定内容边界
                foreach (dynamic para in doc.Paragraphs)
                {
                    try
                    {
                        if (para.Range.Start <= headingEnd) continue; // 跳过当前标题及之前的内容

                        var outlineLevel = para.OutlineLevel;
                        if (outlineLevel != null && outlineLevel >= 1 && outlineLevel <= 9)
                        {
                            int paraLevel = (int)outlineLevel;

                            // 如果是同级或更高级标题，停止
                            if (paraLevel <= headingLevel)
                            {
                                contentEnd = para.Range.Start;
                                Debug.WriteLine($"找到边界标题: {para.Range.Text?.Trim()} (级别: {paraLevel})");
                                break;
                            }

                            // 如果不包含子标题，遇到下级标题也要停止
                            if (!includeSubHeadings && paraLevel > headingLevel)
                            {
                                contentEnd = para.Range.Start;
                                Debug.WriteLine($"不包含子标题，遇到下级标题停止: {para.Range.Text?.Trim()} (级别: {paraLevel})");
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"处理段落边界时出错: {ex.Message}");
                        continue;
                    }
                }

                // 提取指定范围的内容
                if (contentEnd > contentStart)
                {
                    try
                    {
                        var contentRange = doc.Range(contentStart, contentEnd);
                        string rawContent = contentRange.Text ?? "";

                        // 清理内容
                        rawContent = rawContent.Replace("\r", "\n")
                                              .Replace("\x07", "") // 清除表格结束符
                                              .Replace("\f", "\n") // 分页符转换为换行
                                              .Trim();

                        // 按行处理，去除空行和多余空白
                        var lines = rawContent.Split('\n')
                                             .Where(line => !string.IsNullOrWhiteSpace(line))
                                             .Select(line => line.Trim())
                                             .ToArray();

                        rawContent = string.Join("\n", lines);

                        // 应用长度限制
                        if (maxLength > 0 && rawContent.Length > maxLength)
                        {
                            rawContent = rawContent.Substring(0, maxLength) + "...[内容已截断]";
                        }

                        Debug.WriteLine($"提取内容完成，长度: {rawContent.Length} 字符");
                        return rawContent;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"提取内容范围时出错: {ex.Message}");
                        return $"无法提取内容: {ex.Message}";
                    }
                }
                else
                {
                    Debug.WriteLine("标题下没有内容");
                    return "该标题下没有内容。";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"提取标题内容时出错: {ex.Message}");
                return $"提取内容失败: {ex.Message}";
            }
        }









        // 获取颜色名称
        private static string GetColorName(string color)
        {
            switch (color.ToLower())
            {
                case "red": return "红色";
                case "blue": return "蓝色";
                case "green": return "绿色";
                case "black": return "黑色";
                case "white": return "白色";
                case "yellow": return "黄色";
                case "orange": return "橙色";
                case "purple": return "紫色";
                case "gray": return "灰色";
                case "lightblue": return "浅蓝色";
                case "lightgreen": return "浅绿色";
                case "pink": return "粉色";
                case "lightgray": return "浅灰色";
                default: return color;
            }
        }

        // 提取标题级别
        private static int ExtractHeadingLevel(string style)
        {
            if (string.IsNullOrEmpty(style)) return 0;

            var match = System.Text.RegularExpressions.Regex.Match(style, @"标题\s*(\d+)");
            if (match.Success)
            {
                return int.Parse(match.Groups[1].Value);
            }

            match = System.Text.RegularExpressions.Regex.Match(style, @"Heading\s*(\d+)");
            if (match.Success)
            {
                return int.Parse(match.Groups[1].Value);
            }

            return 0;
        }

        // 检测是否为Markdown内容
        private static bool IsMarkdownContent(string content)
        {
            if (string.IsNullOrWhiteSpace(content)) return false;

            // 检测常见的Markdown语法
            return content.Contains("**") ||        // 粗体
                   content.Contains("*") ||         // 斜体或粗体
                   content.Contains("_") ||         // 斜体或粗体
                   content.Contains("`") ||         // 代码
                   content.Contains("[") ||         // 链接
                   content.Contains("#") ||         // 标题
                   content.Contains(">") ||         // 引用
                   content.Contains("- ") ||        // 列表
                   content.Contains("* ") ||        // 列表
                   content.StartsWith("1. ") ||     // 有序列表
                   content.Contains("| ") ||        // 表格
                   content.Contains("```");         // 代码块
        }

        // 检测内容是否包含表格结构（含 [表格预览] 或 紧凑/标准 Markdown 表格）
        private static bool ContainsTableStructure(string content)
        {
            if (string.IsNullOrWhiteSpace(content)) return false;
            if (content.Contains("[表格预览]") || content.Contains("[表格内容]")) return true;

            // 只要存在至少一个“|列1|列2|”样式，再配合分隔线或出现3个以上竖线，也认为包含表格
            if (content.Contains("|"))
            {
                // 标准分隔线特征
                if (System.Text.RegularExpressions.Regex.IsMatch(content, @"\n\s*\|\s*:?[-]{2,}:?\s*(\|\s*:?[-]{2,}:?\s*)+\|?\s*\n"))
                {
                    return true;
                }
                // 紧凑单行：竖线数量>6 且含 --- 段
                if (content.Count(c => c == '|') >= 6 && content.Contains("---"))
                {
                    return true;
                }
            }
            return false;
        }

        // 插入混合内容（段落 + 表格 + 段落）。支持两种情况：
        // 1) 含 [表格预览] 标记；2) 无标记但包含紧凑或标准表格符号
        private static void InsertMixedContent(dynamic range, string content, int indentLevel)
        {
            try
            {
                Debug.WriteLine("=== 开始拆分混合内容 ===");
                string beforeTable = "";
                string tableContent = "";
                string afterTable = "";

                int markIndex = content.IndexOf("[表格预览]");
                if (markIndex < 0) markIndex = content.IndexOf("[表格内容]");

                if (markIndex >= 0)
                {
                    beforeTable = content.Substring(0, markIndex).Trim();
                    string fromMark = content.Substring(markIndex)
                        .Replace("[表格预览]", "")
                        .Replace("[表格内容]", "").TrimStart();

                    // 收集紧随其后的表格片段：直到遇到明显的非表格行
                    var segs = fromMark.Split(new[] { '\r', '\n' }, StringSplitOptions.None);
                    var tbl = new StringBuilder();
                    int cut = -1;
                    bool hasEmptyLine = false;
                    
                    for (int i = 0; i < segs.Length; i++)
                    {
                        var line = segs[i].Trim();
                        
                        // 遇到空行标记，但继续检查下一行
                        if (string.IsNullOrWhiteSpace(line))
                        {
                            hasEmptyLine = true;
                            continue;
                        }
                        
                        // 如果之前遇到过空行，且当前行不是表格行，说明表格结束
                        if (hasEmptyLine && !line.Contains("|"))
                        {
                            cut = i;
                            break;
                        }
                        
                        // 如果当前行包含竖线，认为是表格行
                        if (line.Contains("|"))
                        {
                            tbl.AppendLine(line);
                            hasEmptyLine = false; // 重置空行标记
                        }
                        else
                        {
                            // 遇到非表格行，表格结束
                            cut = i;
                            break;
                        }
                    }
                    tableContent = tbl.ToString();
                    if (cut >= 0 && cut < segs.Length)
                    {
                        afterTable = string.Join("\n", segs.Skip(cut)).Trim();
                    }
                }
                else
                {
                    // 无标记：从分隔线定位表格，再向左回溯到表头行，向右提取若干数据行
                    var sepPattern = @"\|\s*:?[-]{2,}:?\s*(\|\s*:?[-]{2,}:?\s*)+\|";
                    var sepMatch = System.Text.RegularExpressions.Regex.Match(content, sepPattern);
                    if (sepMatch.Success)
                    {
                        // 在分隔线之前的窗口中，右向匹配最后一条表头行（容忍同一行场景，无需换行）
                        int windowStart = Math.Max(0, sepMatch.Index - 800);
                        string leftWindow = content.Substring(windowStart, sepMatch.Index - windowStart);
                        var headerPattern = new System.Text.RegularExpressions.Regex(@"\|\s*[^\|\n]+\s*(\|\s*[^\|\n]+\s*)+\|", System.Text.RegularExpressions.RegexOptions.RightToLeft);
                        var headerMatch = headerPattern.Match(leftWindow);

                        if (headerMatch.Success)
                        {
                            int headerStart = windowStart + headerMatch.Index;
                            string headerRow = headerMatch.Value.Trim();
                            string sepRow = sepMatch.Value.Trim();
                            beforeTable = content.Substring(0, headerStart).Trim();

                            // 表格主体：header + sep + 其后数据行（同一行或多行）
                            string afterSep = content.Substring(sepMatch.Index + sepMatch.Length).TrimStart();
                            
                            // 逐行提取表格数据行，直到遇到非表格行
                            var sepLines = afterSep.Split(new[] { '\r', '\n' }, StringSplitOptions.None);
                            var rowsBuilder = new StringBuilder();
                            int cutIndex = -1;
                            bool hasEmptyLine = false;
                            
                            for (int i = 0; i < sepLines.Length; i++)
                            {
                                var line = sepLines[i].Trim();
                                
                                // 遇到空行标记，但继续检查下一行
                                if (string.IsNullOrWhiteSpace(line))
                                {
                                    hasEmptyLine = true;
                                    continue;
                                }
                                
                                // 如果之前遇到过空行，且当前行不是表格行，说明表格结束
                                if (hasEmptyLine && !line.Contains("|"))
                                {
                                    cutIndex = i;
                                    break;
                                }
                                
                                // 如果当前行包含竖线，认为是表格行
                                if (line.Contains("|"))
                                {
                                    rowsBuilder.AppendLine(line);
                                    hasEmptyLine = false; // 重置空行标记
                                }
                                else
                                {
                                    // 遇到非表格行，表格结束
                                    cutIndex = i;
                                    break;
                                }
                            }
                            
                            string rowsChunk = rowsBuilder.ToString();
                            afterTable = "";
                            if (cutIndex >= 0 && cutIndex < sepLines.Length)
                            {
                                afterTable = string.Join("\n", sepLines.Skip(cutIndex)).Trim();
                            }

                            string compact = headerRow + "\n" + sepRow + "\n" + rowsChunk;
                            tableContent = ConvertCompactTableToMultiline(compact);
                        }
                        else
                        {
                            // 回退策略：使用简化匹配（兼容异常输入）
                            var m = System.Text.RegularExpressions.Regex.Match(content, @"\|[^\n]*\|[^\n]*---[^\n]*\|");
                            if (m.Success)
                            {
                                int start = content.LastIndexOf('|', m.Index);
                                if (start < 0) start = m.Index;
                                beforeTable = content.Substring(0, start).Trim();
                                string from = content.Substring(start);
                                var lines = from.Split(new[] { '\r', '\n' }, StringSplitOptions.None);
                                var tbl = new StringBuilder();
                                int cut = -1;
                                bool hasEmptyLine = false;
                                
                                for (int i = 0; i < lines.Length; i++)
                                {
                                    var line = lines[i].Trim();
                                    
                                    // 遇到空行标记，但继续检查下一行
                                    if (string.IsNullOrWhiteSpace(line))
                                    {
                                        hasEmptyLine = true;
                                        continue;
                                    }
                                    
                                    // 如果之前遇到过空行，且当前行不是表格行，说明表格结束
                                    if (hasEmptyLine && !line.Contains("|"))
                                    {
                                        cut = i;
                                        break;
                                    }
                                    
                                    // 如果当前行包含竖线，认为是表格行
                                    if (line.Contains("|"))
                                    {
                                        tbl.AppendLine(line);
                                        hasEmptyLine = false; // 重置空行标记
                                    }
                                    else
                                    {
                                        // 遇到非表格行，表格结束
                                        cut = i;
                                        break;
                                    }
                                }
                                tableContent = tbl.ToString();
                                if (cut >= 0 && cut < lines.Length)
                                {
                                    afterTable = string.Join("\n", lines.Skip(cut)).Trim();
                                }
                            }
                        }
                    }
                }

                // 若未提取到表格，则回退为纯段落
                if (string.IsNullOrWhiteSpace(tableContent))
                {
                    InsertParagraphContent(range, content, indentLevel);
                    return;
                }

                // 尝试把紧凑单行表格转换为标准多行
                tableContent = ConvertCompactTableToMultiline(tableContent);

                if (!string.IsNullOrWhiteSpace(beforeTable))
                {
                    InsertParagraphContent(range, beforeTable, indentLevel);
                    try
                    {
                        // 将range移动到刚插入文本之后，再插入分隔换行
                        dynamic wordAppTmp = range.Application;
                        dynamic selTmp = wordAppTmp.Selection;
                        range.SetRange(selTmp.Range.End, selTmp.Range.End);
                    }
                    catch { }
                    range.InsertAfter("\r\n\r\n");
                    range.Collapse(0);
                }

                InsertTableContent(range, tableContent, indentLevel);
                try
                {
                    if (range.Tables.Count > 0)
                    {
                        dynamic lastTable = range.Tables[range.Tables.Count];
                        range.SetRange(lastTable.Range.End, lastTable.Range.End);
                    }
                }
                catch { range.Collapse(0); }

                if (!string.IsNullOrWhiteSpace(afterTable))
                {
                    range.InsertAfter("\r\n\r\n");
                    range.Collapse(0);
                    
                    // 递归处理：如果afterTable中还包含表格，继续使用InsertMixedContent
                    if (ContainsTableStructure(afterTable))
                    {
                        Debug.WriteLine("检测到afterTable中还包含表格，递归处理");
                        InsertMixedContent(range, afterTable, indentLevel);
                    }
                    else
                    {
                        InsertParagraphContent(range, afterTable, indentLevel);
                    }
                }

                Debug.WriteLine("=== 混合内容插入完成 ===");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"拆分混合内容时出错: {ex.Message}");
                InsertParagraphContent(range, content, indentLevel);
            }
        }

        // 将紧凑表格转换为标准多行Markdown
        private static string ConvertCompactTableToMultiline(string tableText)
        {
            try
            {
                var lines = tableText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Where(l => l.Trim().Contains("|"))
                                      .ToArray();
                if (lines.Length >= 3) return tableText; // 已是标准表格

                var cells = tableText.Split('|')
                                      .Select(c => c.Trim())
                                      .Where(c => !string.IsNullOrEmpty(c))
                                      .ToList();
                int sepIdx = cells.FindIndex(c => c.All(ch => ch == '-' || ch == ':' || ch == ' '));
                if (sepIdx <= 0) return tableText;
                int col = sepIdx;

                var sb = new StringBuilder();
                var header = cells.Take(col).ToArray();
                sb.AppendLine("| " + string.Join(" | ", header) + " |");
                sb.AppendLine("|" + string.Join("|", Enumerable.Repeat("---", col)) + "|");
                var data = cells.Skip(col + 1).ToArray();
                for (int i = 0; i < data.Length; i += col)
                {
                    var row = data.Skip(i).Take(col).ToArray();
                    if (row.Length == 0) break;
                    if (row.Length < col)
                    {
                        var padded = new string[col];
                        Array.Copy(row, padded, row.Length);
                        for (int j = row.Length; j < col; j++) padded[j] = "";
                        row = padded;
                    }
                    sb.AppendLine("| " + string.Join(" | ", row) + " |");
                }
                return sb.ToString().TrimEnd();
            }
            catch { return tableText; }
        }

        // 将内容转换为Markdown列表格式
        private static string ConvertToMarkdownList(string content)
        {
            if (string.IsNullOrWhiteSpace(content)) return content;

            var lines = content.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            var markdownLines = new List<string>();

            foreach (var line in lines)
            {
                string trimmedLine = line.Trim();
                if (!string.IsNullOrEmpty(trimmedLine))
                {
                    // 如果不是已经格式化的列表项，添加项目符号
                    if (!trimmedLine.StartsWith("- ") &&
                        !trimmedLine.StartsWith("* ") &&
                        !trimmedLine.StartsWith("+ ") &&
                        !System.Text.RegularExpressions.Regex.IsMatch(trimmedLine, @"^\d+\.\s"))
                    {
                        markdownLines.Add("- " + trimmedLine);
                    }
                    else
                    {
                        markdownLines.Add(trimmedLine);
                    }
                }
            }

            return string.Join("\n", markdownLines);
        }

        // 增强的Markdown到HTML转换（支持标题）
        private static string ConvertMarkdownToHtml(string markdown)
        {
            if (string.IsNullOrWhiteSpace(markdown)) return markdown;

            try
            {
                Debug.WriteLine($"开始转换Markdown: {markdown.Substring(0, Math.Min(100, markdown.Length))}...");

                // 基本的Markdown到HTML转换
                string html = markdown;

                // 标题转换（#符号转换为Word标题样式）
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^######\s*(.+)$", "<h6>$1</h6>", System.Text.RegularExpressions.RegexOptions.Multiline);
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^#####\s*(.+)$", "<h5>$1</h5>", System.Text.RegularExpressions.RegexOptions.Multiline);
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^####\s*(.+)$", "<h4>$1</h4>", System.Text.RegularExpressions.RegexOptions.Multiline);
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^###\s*(.+)$", "<h3>$1</h3>", System.Text.RegularExpressions.RegexOptions.Multiline);
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^##\s*(.+)$", "<h2>$1</h2>", System.Text.RegularExpressions.RegexOptions.Multiline);
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^#\s*(.+)$", "<h1>$1</h1>", System.Text.RegularExpressions.RegexOptions.Multiline);

                // 列表转换
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^- (.+)$", "<li>$1</li>", System.Text.RegularExpressions.RegexOptions.Multiline);
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^\* (.+)$", "<li>$1</li>", System.Text.RegularExpressions.RegexOptions.Multiline);
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^\+ (.+)$", "<li>$1</li>", System.Text.RegularExpressions.RegexOptions.Multiline);

                // 有序列表转换
                html = System.Text.RegularExpressions.Regex.Replace(html, @"^(\d+)\.\s(.+)$", "<li>$2</li>", System.Text.RegularExpressions.RegexOptions.Multiline);

                // 如果包含列表项，包装在ul标签中
                if (html.Contains("<li>"))
                {
                    html = "<ul>" + html + "</ul>";
                }

                // 粗体转换
                html = System.Text.RegularExpressions.Regex.Replace(html, @"\*\*(.+?)\*\*", "<strong>$1</strong>");
                html = System.Text.RegularExpressions.Regex.Replace(html, @"__(.+?)__", "<strong>$1</strong>");

                // 斜体转换
                html = System.Text.RegularExpressions.Regex.Replace(html, @"\*(.+?)\*", "<em>$1</em>");
                html = System.Text.RegularExpressions.Regex.Replace(html, @"_(.+?)_", "<em>$1</em>");

                // 代码转换
                html = System.Text.RegularExpressions.Regex.Replace(html, @"`(.+?)`", "<code>$1</code>");

                // 换行转换
                html = html.Replace("\n", "<br>");

                Debug.WriteLine($"Markdown转HTML完成: {html}");
                return html;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Markdown转HTML时出错: {ex.Message}");
                return markdown; // 回退到原始内容
            }
        }

        // 查找标题下的合适插入点
        private static dynamic FindInsertionPointAfterHeading(dynamic doc, dynamic headingRange)
        {
            try
            {
                Debug.WriteLine("开始查找标题下的合适插入位置");

                // 获取标题的结束位置
                int headingEnd = headingRange.End;
                Debug.WriteLine($"标题结束位置: {headingEnd}");

                // 获取标题级别
                int headingLevel = 1;
                try
                {
                    dynamic headingPara = headingRange.Paragraphs[1];
                    string headingStyle = headingPara.Style.ToString();
                    headingLevel = ExtractHeadingLevel(headingStyle);
                    Debug.WriteLine($"标题样式: {headingStyle}, 级别: {headingLevel}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"获取标题样式时出错: {ex.Message}");
                }

                // 从标题结束位置开始，向后查找合适的插入位置
                dynamic searchRange = doc.Range(headingEnd, doc.Content.End);
                dynamic lastGoodPosition = headingRange.Duplicate;
                lastGoodPosition.Collapse(0); // 从标题末尾开始

                // 简化的方法：通过Range查找而不是遍历段落
                try
                {
                    dynamic paragraphs = doc.Paragraphs;
                    int totalParas = paragraphs.Count;
                    Debug.WriteLine($"文档总段落数: {totalParas}");

                    // 从文档的每个段落中找到标题后的内容
                    for (int i = 1; i <= totalParas; i++)
                    {
                        try
                        {
                            dynamic para = paragraphs[i];
                            int paraStart = para.Range.Start;
                            int paraEnd = para.Range.End;

                            // 跳过标题之前的段落
                            if (paraEnd <= headingEnd) continue;

                            // 检查段落样式
                            string paraStyle = "";
                            bool isHeading = false;
                            try
                            {
                                paraStyle = para.Style.ToString();
                                isHeading = IsHeadingStyle(paraStyle);
                            }
                            catch (Exception styleEx)
                            {
                                Debug.WriteLine($"获取段落样式失败: {styleEx.Message}");
                            }

                            // 如果是标题，检查级别
                            if (isHeading)
                            {
                                int paraLevel = ExtractHeadingLevel(paraStyle);
                                if (paraLevel <= headingLevel)
                                {
                                    Debug.WriteLine($"遇到同级或更高级标题 (级别{paraLevel})，停止搜索");
                                    break;
                                }
                                Debug.WriteLine($"遇到下级标题 (级别{paraLevel})，继续搜索");
                            }

                            // 更新最后的内容位置
                            lastGoodPosition = para.Range.Duplicate;
                            lastGoodPosition.Collapse(0); // 移动到段落末尾

                            string paraText = para.Range.Text?.ToString()?.Trim() ?? "";
                            string previewText = paraText.Length > 30 ? paraText.Substring(0, 30) + "..." : paraText;
                            Debug.WriteLine($"找到标题下的内容段落 #{i}: {previewText}");
                        }
                        catch (Exception paraEx)
                        {
                            Debug.WriteLine($"处理段落 #{i} 时出错: {paraEx.Message}");
                        }
                    }
                }
                catch (Exception searchEx)
                {
                    Debug.WriteLine($"搜索段落时出错: {searchEx.Message}");
                }

                Debug.WriteLine($"确定插入位置: {lastGoodPosition.Start}");
                return lastGoodPosition;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"查找插入位置时出错: {ex.Message}");
                Debug.WriteLine($"异常详细信息: {ex}");

                // 回退到标题末尾
                try
                {
                    dynamic fallbackRange = headingRange.Duplicate;
                    fallbackRange.Collapse(0);
                    Debug.WriteLine("使用回退位置：标题末尾");
                    return fallbackRange;
                }
                catch (Exception fallbackEx)
                {
                    Debug.WriteLine($"创建回退位置时也出错: {fallbackEx.Message}");
                    return headingRange; // 最后的回退
                }
            }
        }

        // 智能查找最佳插入位置：在标题下已有内容的末尾
        private static dynamic FindBestInsertionPoint(dynamic doc, dynamic headingRange)
        {
            try
            {
                Debug.WriteLine("开始查找最佳插入位置...");

                // 获取标题的级别（优先使用 OutlineLevel，失败再回退样式名解析）
                dynamic headingPara = headingRange.Paragraphs[1];
                string headingStyle = "";
                try { headingStyle = headingPara.Style?.NameLocal ?? headingPara.Style?.Name ?? headingPara.Style.ToString(); } catch { }
                int headingLevel = 0;
                try
                {
                    var ol = headingPara.OutlineLevel;
                    if (ol != null && ol is int)
                    {
                        headingLevel = (int)ol;
                    }
                }
                catch { }

                if (headingLevel < 1 || headingLevel > 9)
                {
                    try { headingLevel = ExtractHeadingLevel(headingStyle); } catch { headingLevel = 1; }
                }

                Debug.WriteLine($"标题级别(推断): {headingLevel}, 样式: {headingStyle}");

                // 从标题的下一个段落开始查找
                dynamic currentRange = headingRange.Duplicate;
                currentRange.Collapse(0); // 移动到标题末尾

                // 移动到下一个段落
                currentRange.Move(4, 1); // wdParagraph = 4

                // 查找该标题下的所有内容，直到遇到同级或更高级标题
                dynamic lastContentRange = currentRange.Duplicate;
                bool foundContent = false;

                while (currentRange.Start < doc.Content.End)
                {
                    try
                    {
                        dynamic para = currentRange.Paragraphs[1];
                        bool isHeadingPara = false;
                        int paraLevel = 0;
                        // 优先用 OutlineLevel 判断
                        try
                        {
                            var olp = para.OutlineLevel;
                            if (olp != null && olp is int)
                            {
                                paraLevel = (int)olp;
                                isHeadingPara = paraLevel >= 1 && paraLevel <= 9;
                            }
                        }
                        catch { }

                        // 回退：用样式名判断
                        if (!isHeadingPara)
                        {
                            string paraStyleName = "";
                            try { paraStyleName = para.Style?.NameLocal ?? para.Style?.Name ?? para.Style.ToString(); } catch { }
                            if (!string.IsNullOrEmpty(paraStyleName) && IsHeadingStyle(paraStyleName))
                            {
                                isHeadingPara = true;
                                paraLevel = ExtractHeadingLevel(paraStyleName);
                            }
                        }

                        if (isHeadingPara)
                        {
                            // 如果遇到同级或更高级标题，停止查找
                            if (paraLevel <= headingLevel)
                            {
                                Debug.WriteLine($"遇到同级或更高级标题 (级别 {paraLevel})，停止查找");
                                break;
                            }
                        }

                        // 检查段落是否有实际内容（不是空段落）
                        string paraText = para.Range.Text?.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(paraText.Replace("\r", "").Replace("\n", "")))
                        {
                            // 更新最后有内容的位置
                            lastContentRange = para.Range.Duplicate;
                            lastContentRange.Collapse(0); // 移动到段落末尾
                            foundContent = true;
                            Debug.WriteLine($"找到内容段落，更新插入位置到: {lastContentRange.Start}");
                        }

                        // 移动到下一个段落
                        if (currentRange.Move(4, 1) == 0) // wdParagraph = 4
                        {
                            break; // 无法移动，到达文档末尾
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"处理段落时出错: {ex.Message}");
                        // 出错时保守返回标题末尾，避免越界
                        dynamic insertRange = headingRange.Duplicate;
                        insertRange.Collapse(0);
                        return insertRange;
                    }
                }

                if (foundContent)
                {
                    Debug.WriteLine($"找到标题下的已有内容，插入位置: {lastContentRange.Start}");
                    return lastContentRange;
                }
                else
                {
                    // 没有找到内容，直接在标题后插入
                    dynamic insertRange = headingRange.Duplicate;
                    insertRange.Collapse(0); // 移动到标题末尾
                    Debug.WriteLine($"标题下无内容，直接在标题后插入: {insertRange.Start}");
                    return insertRange;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"查找最佳插入位置时出错: {ex.Message}");
                // 出错时回退到简单逻辑
                dynamic fallbackRange = headingRange.Duplicate;
                fallbackRange.Collapse(0);
                return fallbackRange;
            }
        }

        // 通过段落搜索查找标题
        private static bool FindHeadingByParagraphSearch(dynamic doc, string targetHeading, out dynamic range)
        {
            range = null;

            try
            {
                Debug.WriteLine($"开始段落搜索，目标标题: {targetHeading}");

                // 预处理目标标题，提取关键词
                string[] keywords = targetHeading.Split(new char[] { ' ', '　', '\t', ':', '：', '-', '–', '—' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (dynamic para in doc.Paragraphs)
                {
                    try
                    {
                        string paraText = para.Range.Text.Trim();
                        if (string.IsNullOrEmpty(paraText)) continue;

                        Debug.WriteLine($"检查段落: {paraText.Substring(0, Math.Min(50, paraText.Length))}...");

                        // 检查是否包含目标标题（完全匹配）
                        if (paraText.Contains(targetHeading))
                        {
                            Debug.WriteLine($"找到完全匹配的标题: {paraText}");
                            range = para.Range;
                            return true;
                        }

                        // 检查是否包含所有关键词
                        bool containsAllKeywords = true;
                        foreach (string keyword in keywords)
                        {
                            if (!paraText.Contains(keyword))
                            {
                                containsAllKeywords = false;
                                break;
                            }
                        }

                        if (containsAllKeywords && keywords.Length > 0)
                        {
                            Debug.WriteLine($"找到包含所有关键词的标题: {paraText}");
                            range = para.Range;
                            return true;
                        }

                        // 检查标题样式
                        try
                        {
                            string style = para.Style.ToString();
                            if (IsHeadingStyle(style))
                            {
                                // 如果是标题样式，检查是否有部分匹配
                                foreach (string keyword in keywords)
                                {
                                    if (keyword.Length > 2 && paraText.Contains(keyword))
                                    {
                                        Debug.WriteLine($"在标题样式段落中找到关键词匹配: {paraText}");
                                        range = para.Range;
                                        return true;
                                    }
                                }
                            }
                        }
                        catch (Exception styleEx)
                        {
                            Debug.WriteLine($"检查段落样式时出错: {styleEx.Message}");
                        }
                    }
                    catch (Exception paraEx)
                    {
                        Debug.WriteLine($"处理段落时出错: {paraEx.Message}");
                    }
                }

                Debug.WriteLine("段落搜索未找到匹配的标题");
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"段落搜索时出错: {ex.Message}");
                return false;
            }
        }




    }
}