using System;
using System.Collections.Generic;
using System.Linq;
using WordCopilotChat.db;
using WordCopilotChat.models;

namespace WordCopilotChat.services
{
    /// <summary>
    /// 应用设置服务
    /// </summary>
    public class AppSettingsService
    {
        private readonly IFreeSql _freeSql;
        
        // 静态缓存，避免频繁数据库查询
        private static Dictionary<string, string> _settingsCache = new Dictionary<string, string>();
        private static readonly object _lockObject = new object();

        public AppSettingsService()
        {
            _freeSql = FreeSqlDB.Sqlite;
        }

        /// <summary>
        /// 初始化应用设置表
        /// </summary>
        public void InitializeAppSettingsTable()
        {
            try
            {
                // 确保表存在
                _freeSql.CodeFirst.SyncStructure<AppSettings>();
                
                // 初始化默认设置
                InitializeDefaultSettings();
                
                // 加载设置到缓存
                LoadSettingsToCache();
                
                System.Diagnostics.Debug.WriteLine("应用设置表初始化完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"初始化应用设置表失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 初始化默认设置
        /// </summary>
        private void InitializeDefaultSettings()
        {
            try
            {
                var defaultSettings = new List<AppSettings>
                {
                    new AppSettings
                    {
                        SettingKey = "default_temperature",
                        SettingValue = "0.7",
                        Description = "默认温度参数（创造性）",
                        DataType = "double",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    },
                    new AppSettings
                    {
                        SettingKey = "default_max_tokens",
                        SettingValue = "4000",
                        Description = "默认最大令牌数",
                        DataType = "int",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    },
                    new AppSettings
                    {
                        SettingKey = "default_top_p",
                        SettingValue = "0.9",
                        Description = "默认Top-P参数（多样性）",
                        DataType = "double",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    },
                    new AppSettings
                    {
                        SettingKey = "chat_temperature",
                        SettingValue = "0.5",
                        Description = "Chat模式温度参数",
                        DataType = "double",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    },
                    new AppSettings
                    {
                        SettingKey = "chat_max_tokens",
                        SettingValue = "2048",
                        Description = "Chat模式最大令牌数",
                        DataType = "int",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    },
                    new AppSettings
                    {
                        SettingKey = "chat_top_p",
                        SettingValue = "0.8",
                        Description = "Chat模式Top-P参数",
                        DataType = "double",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    },
                    new AppSettings
                    {
                        SettingKey = "agent_temperature",
                        SettingValue = "0.7",
                        Description = "Agent模式温度参数",
                        DataType = "double",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    },
                    new AppSettings
                    {
                        SettingKey = "agent_max_tokens",
                        SettingValue = "8192",
                        Description = "Agent模式最大令牌数",
                        DataType = "int",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    },
                    new AppSettings
                    {
                        SettingKey = "agent_top_p",
                        SettingValue = "0.9",
                        Description = "Agent模式Top-P参数",
                        DataType = "double",
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    }
                };

                foreach (var setting in defaultSettings)
                {
                    // 检查是否已存在
                    var existing = _freeSql.Select<AppSettings>()
                        .Where(s => s.SettingKey == setting.SettingKey)
                        .First();
                        
                    if (existing == null)
                    {
                        _freeSql.Insert(setting).ExecuteAffrows();
                        System.Diagnostics.Debug.WriteLine($"创建默认设置: {setting.SettingKey} = {setting.SettingValue}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"初始化默认设置失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 加载设置到缓存
        /// </summary>
        public void LoadSettingsToCache()
        {
            try
            {
                lock (_lockObject)
                {
                    _settingsCache.Clear();
                    var settings = _freeSql.Select<AppSettings>().ToList();
                    
                    foreach (var setting in settings)
                    {
                        _settingsCache[setting.SettingKey] = setting.SettingValue;
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"已加载 {_settingsCache.Count} 个应用设置到缓存");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载设置到缓存失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取设置值
        /// </summary>
        public string GetSetting(string key, string defaultValue = "")
        {
            lock (_lockObject)
            {
                return _settingsCache.TryGetValue(key, out var value) ? value : defaultValue;
            }
        }

        /// <summary>
        /// 获取整数设置值
        /// </summary>
        public int GetIntSetting(string key, int defaultValue = 0)
        {
            var value = GetSetting(key);
            return int.TryParse(value, out var result) ? result : defaultValue;
        }

        /// <summary>
        /// 获取双精度设置值
        /// </summary>
        public double GetDoubleSetting(string key, double defaultValue = 0.0)
        {
            var value = GetSetting(key);
            return double.TryParse(value, out var result) ? result : defaultValue;
        }

        /// <summary>
        /// 获取布尔设置值
        /// </summary>
        public bool GetBoolSetting(string key, bool defaultValue = false)
        {
            var value = GetSetting(key);
            return bool.TryParse(value, out var result) ? result : defaultValue;
        }

        /// <summary>
        /// 更新设置值
        /// </summary>
        public bool UpdateSetting(string key, string value)
        {
            try
            {
                var setting = _freeSql.Select<AppSettings>()
                    .Where(s => s.SettingKey == key)
                    .First();

                if (setting != null)
                {
                    setting.SettingValue = value;
                    setting.UpdatedAt = DateTime.Now;
                    
                    var updated = _freeSql.Update<AppSettings>()
                        .Set(s => s.SettingValue, value)
                        .Set(s => s.UpdatedAt, DateTime.Now)
                        .Where(s => s.SettingKey == key)
                        .ExecuteAffrows() > 0;

                    if (updated)
                    {
                        lock (_lockObject)
                        {
                            _settingsCache[key] = value;
                        }
                        System.Diagnostics.Debug.WriteLine($"更新设置: {key} = {value}");
                        return true;
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"更新设置失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取所有设置
        /// </summary>
        public List<AppSettings> GetAllSettings()
        {
            try
            {
                return _freeSql.Select<AppSettings>()
                    .OrderBy(s => s.SettingKey)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取所有设置失败: {ex.Message}");
                return new List<AppSettings>();
            }
        }

        /// <summary>
        /// 获取AI参数的默认值
        /// </summary>
        public (double temperature, int maxTokens, double topP) GetDefaultAIParameters()
        {
            return (
                GetDoubleSetting("default_temperature", 0.7),
                GetIntSetting("default_max_tokens", 4000),
                GetDoubleSetting("default_top_p", 0.9)
            );
        }

        /// <summary>
        /// 根据聊天模式获取AI参数
        /// </summary>
        public (double temperature, int maxTokens, double topP) GetAIParametersByMode(string chatMode)
        {
            if (chatMode == "chat-agent")
            {
                return (
                    GetDoubleSetting("agent_temperature", 0.7),
                    GetIntSetting("agent_max_tokens", 8192),
                    GetDoubleSetting("agent_top_p", 0.9)
                );
            }
            else
            {
                return (
                    GetDoubleSetting("chat_temperature", 0.5),
                    GetIntSetting("chat_max_tokens", 2048),
                    GetDoubleSetting("chat_top_p", 0.8)
                );
            }
        }
    }
} 