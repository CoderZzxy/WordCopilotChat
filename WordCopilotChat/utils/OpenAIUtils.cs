using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.IO;

using System.Diagnostics;

namespace WordCopilot.utils
{
    // 工具定义类
    public class Tool
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public Dictionary<string, object> Parameters { get; set; }
        public Func<Dictionary<string, object>, System.Threading.Tasks.Task<string>> ExecuteFunction { get; set; }
    }

    // OpenAI响应模型
    public class OpenAIResponse
    {
        public List<Choice> Choices { get; set; }
    }

    public class Choice
    {
        public Message Message { get; set; }
    }

    public class Message
    {
        public string Role { get; set; }
        public string Content { get; set; }

        [JsonProperty("tool_calls")]
        public List<ToolCall> ToolCalls { get; set; }
    }

    public class ToolCall
    {
        public string Id { get; set; }
        public string Type { get; set; }

        [JsonProperty("function")]
        public ToolFunction Function { get; set; }

        // 兼容性属性，映射到Function.Name
        public string Name => Function?.Name;

        // 兼容性属性，映射到Function.Arguments  
        public object Parameters => Function?.Arguments;
    }

    public class ToolFunction
    {
        public string Name { get; set; }
        public string Arguments { get; set; }
    }

    class OpenAIUtils
    {
        // 工具链支持
        private static readonly List<Tool> _tools = new List<Tool>();

        // 工具预览事件
        public static event Action<JObject> OnToolPreviewReady;

        // 工具调用进度事件
        public static event Action<string> OnToolProgress;

        // 全局日志开关（默认关闭，节省磁盘）：仅当为 true 时才写入 openai_requests / openai_errors 日志
        public static bool EnableLogging { get; set; } = false;

        // 用户数据根目录：C:\Users\<User>\.WordCopilotChat
        private static string GetUserDataRoot()
        {
            try
            {
                var userHome = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                var root = Path.Combine(userHome, ".WordCopilotChat");
                Directory.CreateDirectory(root);
                return root;
            }
            catch
            {
                // 兜底：若获取失败，则退回到当前目录
                return AppDomain.CurrentDomain.BaseDirectory;
            }
        }

        // 保存请求到文件用于Postman调试
        private static void SaveRequestForPostman(string baseUrl, string apiKey, JObject requestBody, string tag)
        {
            try
            {
                if (!EnableLogging) return; // 关闭时不写入请求日志
                // 构建保存路径：<插件目录>/logs/openai_requests/yyyyMMdd/HHmmss-fff_{tag}_{随机}.json
                string pluginDirectory = GetUserDataRoot();
                string dateFolder = DateTime.Now.ToString("yyyyMMdd");
                string logDirectory = Path.Combine(pluginDirectory, "logs", "openai_requests", dateFolder);

                // 创建目录
                Directory.CreateDirectory(logDirectory);

                // 生成文件名
                string timestamp = DateTime.Now.ToString("HHmmss-fff");
                string randomSuffix = Guid.NewGuid().ToString("N").Substring(0, 6);
                string filename = $"{timestamp}_{tag}_{randomSuffix}.json";
                string filepath = Path.Combine(logDirectory, filename);

                // 脱敏处理：隐藏API Key的中间部分，只保留前4位和后4位
                string maskedApiKey = apiKey;
                if (!string.IsNullOrEmpty(apiKey) && apiKey.Length > 8)
                {
                    maskedApiKey = $"{apiKey.Substring(0, 4)}...{apiKey.Substring(apiKey.Length - 4)}";
                }

                // 构建完整的请求信息（包括URL、Headers和Body）
                var fullRequest = new JObject
                {
                    ["url"] = baseUrl,
                    ["method"] = "POST",
                    ["headers"] = new JObject
                    {
                        ["Content-Type"] = "application/json",
                        ["Authorization"] = $"Bearer {maskedApiKey}"
                    },
                    ["body"] = requestBody,
                    ["timestamp"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                    ["tag"] = tag,
                    ["note"] = "⚠️ API Key已脱敏处理，实际使用时请替换为完整的API Key"
                };

                // 保存到文件
                File.WriteAllText(filepath, fullRequest.ToString(Formatting.Indented), Encoding.UTF8);

                Debug.WriteLine($"请求已保存到: {filepath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存请求到文件时出错: {ex.Message}");
                // 不抛出异常，避免影响主流程
            }
        }

        // 触发工具进度事件的公共方法
        public static void NotifyToolProgress(string message)
        {
            OnToolProgress?.Invoke(message);
        }

        // 保存错误响应到日志文件
        private static void SaveErrorToLog(string baseUrl, JObject requestBody, string errorResponse, string statusCode, string tag)
        {
            try
            {
                if (!EnableLogging) return; // 关闭时不写入错误日志
                // 构建保存路径
                string pluginDirectory = GetUserDataRoot();
                string dateFolder = DateTime.Now.ToString("yyyyMMdd");
                string logDirectory = Path.Combine(pluginDirectory, "logs", "openai_errors", dateFolder);

                // 创建目录
                Directory.CreateDirectory(logDirectory);

                // 生成文件名
                string timestamp = DateTime.Now.ToString("HHmmss-fff");
                string randomSuffix = Guid.NewGuid().ToString("N").Substring(0, 6);
                string filename = $"{timestamp}_{tag}_{statusCode}_{randomSuffix}.json";
                string filepath = Path.Combine(logDirectory, filename);

                // 构建错误日志内容
                var errorLog = new JObject
                {
                    ["url"] = baseUrl,
                    ["method"] = "POST",
                    ["statusCode"] = statusCode,
                    ["request"] = requestBody,
                    ["errorResponse"] = errorResponse,
                    ["timestamp"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                    ["tag"] = tag
                };

                // 保存到文件
                File.WriteAllText(filepath, errorLog.ToString(Formatting.Indented), Encoding.UTF8);

                Debug.WriteLine($"错误日志已保存到: {filepath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存错误日志时出错: {ex.Message}");
            }
        }

        // 注册工具
        public static void RegisterTool(string name, string description, Dictionary<string, object> parameters, Func<Dictionary<string, object>, System.Threading.Tasks.Task<string>> executeFunction)
        {
            // 检查是否已经注册过同名工具
            var existingTool = _tools.FirstOrDefault(t => t.Name == name);
            if (existingTool != null)
            {
                Debug.WriteLine($"工具 {name} 已存在，将被替换");
                _tools.Remove(existingTool);
            }

            _tools.Add(new Tool
            {
                Name = name,
                Description = description,
                Parameters = parameters,
                ExecuteFunction = executeFunction
            });

            Debug.WriteLine($"工具 {name} 注册成功");
        }

        // 清空所有工具
        public static void ClearTools()
        {
            _tools.Clear();
            Debug.WriteLine("所有工具已清空");
        }

        // 获取已注册的工具数量
        public static int GetToolCount()
        {
            return _tools.Count;
        }

        // 原有的简单API调用方法（保持向后兼容）
        public static async Task OpenAIApiClientAsync(string baseUrl, string apiKey, string json, CancellationToken cancellationToken, Action<string> onContentReceived)
        {
            await OpenAIApiClientAsync(baseUrl, apiKey, json, cancellationToken, onContentReceived, null);
        }

        // 增强版API调用方法，支持工具链
        public static async Task OpenAIApiClientAsync(string baseUrl, string apiKey, string json, CancellationToken cancellationToken, Action<string> onContentReceived, List<JObject> messages)
        {
            // 创建 HttpClient 实例
            HttpClient _httpClient = new HttpClient(new HttpClientHandler()
            {
                // 支持TLS 1.2和1.3,否则无法正常请求https请求
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13
            });
            // 解析 JSON 字符串（移到 try 外以便在 catch 中使用）
            JObject jsonObject = null;
            try
            {
                // 解析 JSON 字符串
            
                jsonObject = JObject.Parse(json);
                bool stream = jsonObject["stream"] != null && jsonObject["stream"].Value<bool>();

                // 如果有工具且是智能体模式，添加工具定义
                if (_tools.Count > 0 && messages != null)
                {
                    Debug.WriteLine($"检测到 {_tools.Count} 个工具，添加到请求中");

                    // 保存Agent模式的初始请求（在添加工具之前）
                    SaveRequestForPostman(baseUrl, apiKey, jsonObject, "agent_initial");

                    await CallWithTools(baseUrl, apiKey, jsonObject, cancellationToken, onContentReceived, messages, _httpClient);
                    return;
                }

                // 设置超时
                _httpClient.Timeout = TimeSpan.FromMinutes(5);

                // 保存普通模式的请求
                SaveRequestForPostman(baseUrl, apiKey, jsonObject, "general");

                var request = new HttpRequestMessage(HttpMethod.Post, baseUrl)
                {
                    Headers =
            {
                Authorization = new AuthenticationHeaderValue("Bearer", apiKey)
            },
                    Content = new StringContent(JsonConvert.SerializeObject(jsonObject), Encoding.UTF8, "application/json")
                };

                using (var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken))
                {
                    response.EnsureSuccessStatusCode();

                    if (stream)
                    {
                        using (var streamReader = new System.IO.StreamReader(await response.Content.ReadAsStreamAsync()))
                        {
                            string line;
                            while ((line = await streamReader.ReadLineAsync()) != null)
                            {
                                // 检查是否被取消
                                cancellationToken.ThrowIfCancellationRequested();

                                if (!string.IsNullOrWhiteSpace(line))
                                {
                                    var cleanLine = line.Trim();
                                    if (cleanLine.StartsWith("data:"))
                                    {
                                        cleanLine = cleanLine.Substring(5).Trim();
                                    }

                                    try
                                    {
                                        var decodedLine = JsonConvert.DeserializeObject<JObject>(cleanLine);
                                        if (decodedLine?["choices"] is JArray choices)
                                        {
                                            foreach (var choice in choices)
                                            {
                                                if (choice["delta"]?["content"] is JToken content)
                                                {
                                                    var contentString = content.ToString();
                                                    onContentReceived?.Invoke(contentString);
                                                }
                                            }
                                        }
                                    }
                                    catch (JsonException ex)
                                    {
                                        Console.WriteLine("JSON decode error: " + ex.Message);
                                        continue;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        var responseBody = await response.Content.ReadAsStringAsync();
                        var decodedLine = JsonConvert.DeserializeObject<JObject>(responseBody);
                        var content = decodedLine?["choices"]?[0]?["message"]?["content"]?.ToString();
                        Console.WriteLine("Content: " + content);
                        onContentReceived?.Invoke(content);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("API请求被取消");
                throw;
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"Request error: {e.Message}");

                // 记录HTTP请求错误到日志
                try
                {
                    var errorDetail = new
                    {
                        error_type = "HttpRequestException",
                        message = e.Message,
                        stack_trace = e.StackTrace,
                        inner_exception = e.InnerException?.Message
                    };
                    // 如果 jsonObject 为 null（JSON解析失败），创建一个包含原始json的对象
                    var requestBody = jsonObject ?? new JObject { ["raw_json"] = json };
                    SaveErrorToLog(baseUrl, requestBody, JsonConvert.SerializeObject(errorDetail), "HttpError", "general_http_error");
                }
                catch (Exception logEx)
                {
                    Debug.WriteLine($"保存HTTP错误日志失败: {logEx.Message}");
                }

                MessageBox.Show($"Request error: {e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unexpected error: {e.Message}");

                // 记录未预期错误到日志
                try
                {
                    var errorDetail = new
                    {
                        error_type = e.GetType().Name,
                        message = e.Message,
                        stack_trace = e.StackTrace,
                        inner_exception = e.InnerException?.Message
                    };
                    SaveErrorToLog(baseUrl, jsonObject, JsonConvert.SerializeObject(errorDetail), "UnexpectedError", "general_unexpected_error");
                }
                catch (Exception logEx)
                {
                    Debug.WriteLine($"保存未预期错误日志失败: {logEx.Message}");
                }

                MessageBox.Show($"Unexpected error: {e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _httpClient?.Dispose();
            }
        }

        // 支持工具链的API调用
        private static async Task CallWithTools(string baseUrl, string apiKey, JObject requestBody, CancellationToken cancellationToken, Action<string> onContentReceived, List<JObject> conversationMessages, HttpClient httpClient)
        {
            try
            {
                // 设置超时
                httpClient.Timeout = TimeSpan.FromMinutes(5);

                // 准备工具定义
                var functionDefinitions = new JArray();
                foreach (var tool in _tools)
                {
                    functionDefinitions.Add(new JObject
                    {
                        ["type"] = "function",
                        ["function"] = new JObject
                        {
                            ["name"] = tool.Name,
                            ["description"] = tool.Description,
                            ["parameters"] = new JObject
                            {
                                ["type"] = "object",
                                ["properties"] = JObject.FromObject(tool.Parameters),
                                ["required"] = new JArray() // 可以根据需要设置必填参数
                            }
                        }
                    });
                }

                // 添加工具到请求中
                requestBody["tools"] = functionDefinitions;
                requestBody["tool_choice"] = "auto";

                // 支持流式工具调用（现代AI工具的标准做法）
                requestBody["stream"] = true;

                Debug.WriteLine($"发送带工具的请求，工具数量: {functionDefinitions.Count}");
                Debug.WriteLine("使用流式响应模式支持现代AI工具体验");

                // 输出完整的请求JSON用于调试
                Debug.WriteLine("=== 完整请求JSON ===");
                Debug.WriteLine(requestBody.ToString(Formatting.Indented));
                Debug.WriteLine("=== 请求JSON结束 ===");

                // 使用流式调用支持工具链
                await CallWithToolsStreaming(baseUrl, apiKey, requestBody, httpClient, cancellationToken, onContentReceived, conversationMessages);


            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("工具链API请求被取消");
                throw;
            }
            catch (JsonException jsonEx)
            {
                Debug.WriteLine($"工具链JSON解析错误: {jsonEx.Message}");

                // 记录JSON解析错误到日志
                try
                {
                    var errorDetail = new
                    {
                        error_type = "JsonException",
                        message = jsonEx.Message,
                        stack_trace = jsonEx.StackTrace,
                        inner_exception = jsonEx.InnerException?.Message
                    };
                    SaveErrorToLog(baseUrl, requestBody, JsonConvert.SerializeObject(errorDetail), "JsonError", "agent_json_error");
                }
                catch (Exception logEx)
                {
                    Debug.WriteLine($"保存JSON错误日志失败: {logEx.Message}");
                }

                // 可能是API返回了HTML错误页面，提供更友好的错误信息
                onContentReceived?.Invoke($"⚠️ API调用失败：服务器返回了无效的响应格式。请检查API密钥是否正确，或者API服务是否正常。\n\n错误详情：{jsonEx.Message}");
            }
            catch (HttpRequestException httpEx)
            {
                Debug.WriteLine($"工具链HTTP请求错误: {httpEx.Message}");

                // 记录HTTP请求错误到日志
                try
                {
                    var errorDetail = new
                    {
                        error_type = "HttpRequestException",
                        message = httpEx.Message,
                        stack_trace = httpEx.StackTrace,
                        inner_exception = httpEx.InnerException?.Message
                    };
                    SaveErrorToLog(baseUrl, requestBody, JsonConvert.SerializeObject(errorDetail), "HttpError", "agent_http_error");
                }
                catch (Exception logEx)
                {
                    Debug.WriteLine($"保存HTTP错误日志失败: {logEx.Message}");
                }

                onContentReceived?.Invoke($"⚠️ 网络请求失败：{httpEx.Message}\n\n请检查网络连接和API服务状态。");
            }
            catch (TimeoutException timeoutEx)
            {
                Debug.WriteLine($"工具链请求超时: {timeoutEx.Message}");

                // 记录超时错误到日志
                try
                {
                    var errorDetail = new
                    {
                        error_type = "TimeoutException",
                        message = timeoutEx.Message,
                        stack_trace = timeoutEx.StackTrace
                    };
                    SaveErrorToLog(baseUrl, requestBody, JsonConvert.SerializeObject(errorDetail), "Timeout", "agent_timeout_error");
                }
                catch (Exception logEx)
                {
                    Debug.WriteLine($"保存超时错误日志失败: {logEx.Message}");
                }

                onContentReceived?.Invoke("⚠️ API请求超时，请稍后重试。");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"工具链调用出错: {ex.Message}");
                Debug.WriteLine($"错误堆栈: {ex.StackTrace}");

                // 记录未预期错误到日志
                try
                {
                    var errorDetail = new
                    {
                        error_type = ex.GetType().Name,
                        message = ex.Message,
                        stack_trace = ex.StackTrace,
                        inner_exception = ex.InnerException?.Message
                    };
                    SaveErrorToLog(baseUrl, requestBody, JsonConvert.SerializeObject(errorDetail), "UnexpectedError", "agent_unexpected_error");
                }
                catch (Exception logEx)
                {
                    Debug.WriteLine($"保存未预期错误日志失败: {logEx.Message}");
                }

                onContentReceived?.Invoke($"⚠️ 工具链调用失败：{ex.Message}\n\n已回退到普通聊天模式。");
            }
        }

        // 调用OpenAI API的辅助方法
        private static async Task<OpenAIResponse> CallOpenAI(string baseUrl, string apiKey, JObject requestBody, HttpClient httpClient, CancellationToken cancellationToken)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, baseUrl)
            {
                Headers =
                {
                    Authorization = new AuthenticationHeaderValue("Bearer", apiKey)
                },
                Content = new StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
            };

            try
            {
                Debug.WriteLine($"发送工具链API请求到: {baseUrl}");
                Debug.WriteLine($"请求内容: {requestBody.ToString().Substring(0, Math.Min(200, requestBody.ToString().Length))}...");

                using (var response = await httpClient.SendAsync(request, cancellationToken))
                {
                    var responseJson = await response.Content.ReadAsStringAsync();
                    Debug.WriteLine($"API响应状态: {response.StatusCode}");
                    Debug.WriteLine($"响应内容类型: {response.Content.Headers.ContentType}");
                    Debug.WriteLine($"响应内容: {responseJson.Substring(0, Math.Min(500, responseJson.Length))}...");

                    // 检查响应状态码
                    if (!response.IsSuccessStatusCode)
                    {
                        Debug.WriteLine($"API返回错误状态码: {response.StatusCode}");
                        Debug.WriteLine($"错误响应内容: {responseJson}");

                        // 尝试解析错误详情
                        string errorMessage = $"响应状态代码不指示成功: {(int)response.StatusCode} ({response.ReasonPhrase})";
                        try
                        {
                            var errorJson = JObject.Parse(responseJson);
                            var errorDetail = errorJson["error"]?["message"]?.ToString();
                            var errorCode = errorJson["error"]?["code"]?.ToString();

                            if (!string.IsNullOrEmpty(errorDetail))
                            {
                                errorMessage = errorDetail;

                                // 特殊处理常见错误类型
                                if (errorMessage.Contains("max_tokens") || errorMessage.Contains("maximum context length") || errorMessage.Contains("token"))
                                {
                                    errorMessage = $"⚠️ Token设置错误：{errorDetail}\n\n建议：请在设置中降低Max Tokens值，或减少对话历史长度。";
                                }
                                else if (!string.IsNullOrEmpty(errorCode))
                                {
                                    errorMessage = $"错误代码 {errorCode}：{errorDetail}";
                                }
                            }

                            Debug.WriteLine($"解析后的错误信息: {errorMessage}");

                            // 保存错误响应到日志
                            try
                            {
                                SaveErrorToLog(baseUrl, requestBody, responseJson, response.StatusCode.ToString(), "general_error");
                            }
                            catch (Exception logEx)
                            {
                                Debug.WriteLine($"保存错误日志失败: {logEx.Message}");
                            }
                        }
                        catch (Exception parseEx)
                        {
                            Debug.WriteLine($"解析错误响应失败: {parseEx.Message}");
                        }

                        throw new HttpRequestException(errorMessage);
                    }

                    // 检查是否为SSE流式响应格式
                    if (response.Content.Headers.ContentType?.MediaType == "text/event-stream")
                    {
                        Debug.WriteLine("检测到SSE流式响应，尝试提取JSON内容");
                        responseJson = ExtractJsonFromSSE(responseJson);
                        Debug.WriteLine($"提取后的JSON: {responseJson.Substring(0, Math.Min(200, responseJson.Length))}...");
                    }

                    // 检查响应是否为JSON格式
                    if (!responseJson.TrimStart().StartsWith("{") && !responseJson.TrimStart().StartsWith("["))
                    {
                        Debug.WriteLine("API返回的不是JSON格式的响应");
                        throw new JsonException($"API返回非JSON格式响应: {responseJson.Substring(0, Math.Min(100, responseJson.Length))}");
                    }

                    try
                    {
                        var result = JsonConvert.DeserializeObject<OpenAIResponse>(responseJson);
                        Debug.WriteLine("JSON解析成功");
                        return result;
                    }
                    catch (JsonException jsonEx)
                    {
                        Debug.WriteLine($"JSON解析失败: {jsonEx.Message}");
                        Debug.WriteLine($"响应原文: {responseJson}");
                        throw new JsonException($"解析API响应失败: {jsonEx.Message}. 响应内容: {responseJson.Substring(0, Math.Min(200, responseJson.Length))}");
                    }
                }
            }
            catch (HttpRequestException httpEx)
            {
                Debug.WriteLine($"HTTP请求异常: {httpEx.Message}");
                throw;
            }
            catch (TaskCanceledException)
            {
                Debug.WriteLine("API请求超时或被取消");
                throw new TimeoutException("API请求超时");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CallOpenAI发生未知错误: {ex.Message}");
                throw;
            }
        }

        // 从SSE流式响应中提取JSON内容
        private static string ExtractJsonFromSSE(string sseResponse)
        {
            try
            {
                Debug.WriteLine("开始从SSE响应提取JSON内容");

                // 分割SSE数据行
                var lines = sseResponse.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                var jsonParts = new List<JObject>();

                foreach (var line in lines)
                {
                    if (line.StartsWith("data: ") && !line.Contains("[DONE]"))
                    {
                        var jsonContent = line.Substring(6).Trim(); // 移除 "data: " 前缀
                        try
                        {
                            var jsonObj = JObject.Parse(jsonContent);
                            jsonParts.Add(jsonObj);
                        }
                        catch (JsonException)
                        {
                            Debug.WriteLine($"无法解析JSON片段: {jsonContent.Substring(0, Math.Min(50, jsonContent.Length))}");
                        }
                    }
                }

                if (jsonParts.Count == 0)
                {
                    throw new JsonException("未能从SSE响应中提取有效的JSON数据");
                }

                // 合并tool_calls和内容
                var result = new JObject
                {
                    ["choices"] = new JArray
                    {
                        new JObject
                        {
                            ["message"] = new JObject
                            {
                                ["role"] = "assistant",
                                ["content"] = "",
                                ["tool_calls"] = new JArray()
                            }
                        }
                    }
                };

                var toolCallsMap = new Dictionary<int, JObject>(); // 使用索引来合并分段的tool_calls
                var contentParts = new List<string>();

                foreach (var part in jsonParts)
                {
                    var choices = part["choices"] as JArray;
                    if (choices != null && choices.Count > 0)
                    {
                        var delta = choices[0]["delta"];
                        if (delta != null)
                        {
                            // 提取content
                            var content = delta["content"]?.ToString();
                            if (!string.IsNullOrEmpty(content))
                            {
                                contentParts.Add(content);
                            }

                            // 提取tool_calls
                            var deltaToolCalls = delta["tool_calls"] as JArray;
                            if (deltaToolCalls != null)
                            {
                                for (int i = 0; i < deltaToolCalls.Count; i++)
                                {
                                    var toolCallDelta = deltaToolCalls[i] as JObject;
                                    if (toolCallDelta != null)
                                    {
                                        var index = toolCallDelta["index"]?.ToObject<int>() ?? i;

                                        if (!toolCallsMap.ContainsKey(index))
                                        {
                                            toolCallsMap[index] = new JObject();
                                        }

                                        var existingToolCall = toolCallsMap[index];

                                        // 合并id
                                        if (toolCallDelta["id"] != null)
                                        {
                                            existingToolCall["id"] = toolCallDelta["id"];
                                        }

                                        // 合并type
                                        if (toolCallDelta["type"] != null)
                                        {
                                            existingToolCall["type"] = toolCallDelta["type"];
                                        }

                                        // 合并function
                                        if (toolCallDelta["function"] != null)
                                        {
                                            var functionDelta = toolCallDelta["function"] as JObject;
                                            if (existingToolCall["function"] == null)
                                            {
                                                existingToolCall["function"] = new JObject();
                                            }

                                            var existingFunction = existingToolCall["function"] as JObject;

                                            if (functionDelta["name"] != null)
                                            {
                                                existingFunction["name"] = functionDelta["name"];
                                            }

                                            if (functionDelta["arguments"] != null)
                                            {
                                                var args = functionDelta["arguments"].ToString();
                                                if (existingFunction["arguments"] == null)
                                                {
                                                    existingFunction["arguments"] = args;
                                                }
                                                else
                                                {
                                                    existingFunction["arguments"] = existingFunction["arguments"].ToString() + args;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 设置合并后的内容
                result["choices"][0]["message"]["content"] = string.Join("", contentParts);

                // 设置tool_calls
                if (toolCallsMap.Count > 0)
                {
                    var mergedToolCalls = new JArray();
                    foreach (var toolCall in toolCallsMap.Values)
                    {
                        mergedToolCalls.Add(toolCall);
                    }
                    result["choices"][0]["message"]["tool_calls"] = mergedToolCalls;
                }

                Debug.WriteLine($"SSE提取完成，工具调用数量: {toolCallsMap.Count}");
                return result.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SSE提取失败: {ex.Message}");
                throw new JsonException($"无法从SSE响应提取JSON: {ex.Message}");
            }
        }

        // 流式工具调用方法 - 支持现代AI工具体验
        private static async Task CallWithToolsStreaming(string baseUrl, string apiKey, JObject requestBody, HttpClient httpClient, CancellationToken cancellationToken, Action<string> onContentReceived, List<JObject> conversationMessages)
        {
            Debug.WriteLine("CallWithToolsStreaming: 开始调用");

            var request = new HttpRequestMessage(HttpMethod.Post, baseUrl)
            {
                Headers = { Authorization = new AuthenticationHeaderValue("Bearer", apiKey) },
                Content = new StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
            };

            using (var response = await httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken))
            {
                // 检查响应状态码
                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    Debug.WriteLine($"API返回错误状态码: {response.StatusCode}");
                    Debug.WriteLine($"错误响应内容: {errorContent}");

                    // 尝试解析错误详情
                    string errorMessage = $"响应状态代码不指示成功: {(int)response.StatusCode} ({response.ReasonPhrase})";
                    try
                    {
                        var errorJson = JObject.Parse(errorContent);
                        var errorDetail = errorJson["error"]?["message"]?.ToString();
                        var errorCode = errorJson["error"]?["code"]?.ToString();
                        var errorType = errorJson["error"]?["type"]?.ToString();

                        if (!string.IsNullOrEmpty(errorDetail))
                        {
                            errorMessage = errorDetail;

                            // 特殊处理常见错误类型
                            if (errorMessage.Contains("max_tokens") || errorMessage.Contains("maximum context length") || errorMessage.Contains("token"))
                            {
                                errorMessage = $"⚠️ Token设置错误：{errorDetail}\n\n建议：请在设置中降低Max Tokens值，或减少对话历史长度。";
                            }
                            else if (!string.IsNullOrEmpty(errorCode))
                            {
                                errorMessage = $"错误代码 {errorCode}：{errorDetail}";
                            }
                        }

                        Debug.WriteLine($"解析后的错误信息: {errorMessage}");

                        // 保存错误响应到日志
                        try
                        {
                            SaveErrorToLog(baseUrl, requestBody, errorContent, response.StatusCode.ToString(), "agent_error");
                        }
                        catch (Exception logEx)
                        {
                            Debug.WriteLine($"保存错误日志失败: {logEx.Message}");
                        }
                    }
                    catch (Exception parseEx)
                    {
                        Debug.WriteLine($"解析错误响应失败: {parseEx.Message}");
                    }

                    throw new HttpRequestException(errorMessage);
                }

                using (var stream = await response.Content.ReadAsStreamAsync())
                using (var reader = new StreamReader(stream))
                {
                    var currentToolCalls = new List<dynamic>();
                    var currentContent = new StringBuilder();

                    string line;
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            break;

                        if (line.StartsWith("data: "))
                        {
                            var jsonData = line.Substring(6);
                            if (jsonData == "[DONE]")
                                break;

                            try
                            {
                                var chunk = JObject.Parse(jsonData);
                                var choices = chunk["choices"] as JArray;

                                if (choices != null && choices.Count > 0)
                                {
                                    var choice = choices[0] as JObject;
                                    var delta = choice["delta"] as JObject;

                                    if (delta != null)
                                    {
                                        // 处理文本内容 - 实现内联效果
                                        var content = delta["content"]?.ToString();
                                        if (!string.IsNullOrEmpty(content))
                                        {
                                            currentContent.Append(content);
                                            onContentReceived?.Invoke(content);
                                        }

                                        // 处理工具调用
                                        var toolCalls = delta["tool_calls"] as JArray;
                                        if (toolCalls != null)
                                        {
                                            foreach (var toolCallDelta in toolCalls)
                                            {
                                                var index = toolCallDelta["index"]?.ToObject<int>() ?? 0;
                                                var id = toolCallDelta["id"]?.ToString();
                                                var function = toolCallDelta["function"] as JObject;

                                                // 确保工具调用列表足够大
                                                while (currentToolCalls.Count <= index)
                                                {
                                                    currentToolCalls.Add(new { id = "", name = "", arguments = new StringBuilder() });
                                                }

                                                // 更新工具调用信息
                                                if (!string.IsNullOrEmpty(id))
                                                {
                                                    currentToolCalls[index] = new
                                                    {
                                                        id = id,
                                                        name = currentToolCalls[index].name,
                                                        arguments = currentToolCalls[index].arguments
                                                    };
                                                }

                                                if (function != null)
                                                {
                                                    var name = function["name"]?.ToString();
                                                    var arguments = function["arguments"]?.ToString();

                                                    if (!string.IsNullOrEmpty(name))
                                                    {
                                                        currentToolCalls[index] = new
                                                        {
                                                            id = currentToolCalls[index].id,
                                                            name = name,
                                                            arguments = currentToolCalls[index].arguments
                                                        };

                                                        // 发送工具调用开始进度
                                                        OnToolProgress?.Invoke($"执行工具: {name}");
                                                    }

                                                    if (!string.IsNullOrEmpty(arguments))
                                                    {
                                                        ((StringBuilder)currentToolCalls[index].arguments).Append(arguments);
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    // 检查是否完成
                                    var finishReason = choice["finish_reason"]?.ToString();
                                    if (finishReason == "tool_calls")
                                    {
                                        Debug.WriteLine("CallWithToolsStreaming: 检测到tool_calls，开始执行工具");
                                        // 执行工具调用并继续对话
                                        await ExecuteToolCallsAsync(currentToolCalls, conversationMessages, baseUrl, apiKey, requestBody, httpClient, cancellationToken, onContentReceived);
                                        return;
                                    }
                                    else if (!string.IsNullOrEmpty(finishReason))
                                    {
                                        Debug.WriteLine($"CallWithToolsStreaming: 检测到完成原因: {finishReason}");
                                    }
                                }
                            }
                            catch (JsonException ex)
                            {
                                Debug.WriteLine($"解析流式响应时出错: {ex.Message}");
                            }
                        }
                    }

                    Debug.WriteLine($"CallWithToolsStreaming: 流式处理结束，内容长度: {currentContent.Length}");
                }
            }

            Debug.WriteLine("CallWithToolsStreaming: 方法结束");
        }

        // 执行工具调用的辅助方法
        private static async Task ExecuteToolCallsAsync(List<dynamic> toolCalls, List<JObject> conversationMessages, string baseUrl, string apiKey, JObject requestBody, HttpClient httpClient, CancellationToken cancellationToken, Action<string> onContentReceived)
        {
            // 添加助手消息到对话历史
            var toolCallsArray = new JArray();
            foreach (var toolCall in toolCalls)
            {
                toolCallsArray.Add(new JObject
                {
                    ["id"] = toolCall.id,
                    ["type"] = "function",
                    ["function"] = new JObject
                    {
                        ["name"] = toolCall.name,
                        ["arguments"] = toolCall.arguments.ToString()
                    }
                });
            }

            conversationMessages.Add(new JObject
            {
                ["role"] = "assistant",
                ["content"] = "",
                ["tool_calls"] = toolCallsArray
            });

            // 执行每个工具调用
            foreach (var toolCall in toolCalls)
            {
                var tool = _tools.FirstOrDefault(t => t.Name == toolCall.name);
                if (tool != null)
                {
                    try
                    {
                        var parameters = JsonConvert.DeserializeObject<Dictionary<string, object>>(toolCall.arguments.ToString());
                        var output = await tool.ExecuteFunction(parameters);

                        // 检查是否是预览模式
                        JObject outputJson = null;
                        try { outputJson = JObject.Parse(output); } catch { }

                        if (outputJson != null && outputJson["preview_mode"]?.ToObject<bool>() == true && outputJson["success"]?.ToObject<bool>() == true)
                        {
                            OnToolPreviewReady?.Invoke(outputJson);
                            conversationMessages.Add(new JObject
                            {
                                ["role"] = "tool",
                                ["tool_call_id"] = toolCall.id,
                                ["content"] = "预览操作已生成，等待用户确认。"
                            });

                            // 预览模式下不继续调用API，直接返回等待用户操作
                            Debug.WriteLine("预览模式：暂停API调用，等待用户确认操作");
                            return;
                        }
                        else
                        {
                            conversationMessages.Add(new JObject
                            {
                                ["role"] = "tool",
                                ["tool_call_id"] = toolCall.id,
                                ["content"] = output
                            });

                            OnToolProgress?.Invoke($"工具 {toolCall.name} 执行完成，返回数据长度: {output?.Length ?? 0} 字符");
                        }
                    }
                    catch (Exception ex)
                    {
                        conversationMessages.Add(new JObject
                        {
                            ["role"] = "tool",
                            ["tool_call_id"] = toolCall.id,
                            ["content"] = $"工具执行失败: {ex.Message}"
                        });
                    }
                }
            }

            // 继续对话 - 递归调用实现多轮工具调用
            requestBody["messages"] = JArray.FromObject(conversationMessages);

            // 保存Agent模式的后续请求（工具执行后）
            SaveRequestForPostman(baseUrl, apiKey, requestBody, "agent_followup");

            Debug.WriteLine($"ExecuteToolCallsAsync: 工具执行完成，继续调用API获取最终回复，对话历史数量: {conversationMessages.Count}");

            // 创建新的 HttpClient 实例用于递归调用，避免连接池问题
            using (var newHttpClient = new HttpClient(new HttpClientHandler()
            {
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13
            }))
            {
                newHttpClient.Timeout = TimeSpan.FromMinutes(5);
                Debug.WriteLine("ExecuteToolCallsAsync: 创建新的HttpClient实例用于递归调用");

                await CallWithToolsStreaming(baseUrl, apiKey, requestBody, newHttpClient, cancellationToken, onContentReceived, conversationMessages);

                Debug.WriteLine("ExecuteToolCallsAsync: 递归调用完成");
            }
        }

        // 保持向后兼容的重载方法
        public static async Task OpenAIApiClientAsync(string baseUrl, string apiKey, string json, Action<string> onContentReceived)
        {
            await OpenAIApiClientAsync(baseUrl, apiKey, json, CancellationToken.None, onContentReceived);
        }
    }
}
