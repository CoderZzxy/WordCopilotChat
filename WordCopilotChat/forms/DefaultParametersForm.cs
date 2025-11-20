using System;
using System.Drawing;
using System.Windows.Forms;
using WordCopilotChat.services;

namespace WordCopilotChat
{
    public partial class DefaultParametersForm : Form
    {
        private AppSettingsService _appSettingsService;

        public DefaultParametersForm(AppSettingsService appSettingsService)
        {
            _appSettingsService = appSettingsService;
            InitializeComponent();
            LoadCurrentSettings();
        }

        private void LoadCurrentSettings()
        {
            try
            {
                // 加载通用默认参数
                nudDefaultTemp.Value = (decimal)_appSettingsService.GetDoubleSetting("default_temperature", 0.7);
                nudDefaultMaxTokens.Value = _appSettingsService.GetIntSetting("default_max_tokens", 4000);
                nudDefaultTopP.Value = (decimal)_appSettingsService.GetDoubleSetting("default_top_p", 0.9);
                
                // 加载Chat模式参数
                nudChatTemp.Value = (decimal)_appSettingsService.GetDoubleSetting("chat_temperature", 0.5);
                nudChatMaxTokens.Value = _appSettingsService.GetIntSetting("chat_max_tokens", 2048);
                nudChatTopP.Value = (decimal)_appSettingsService.GetDoubleSetting("chat_top_p", 0.8);
                
                // 加载Agent模式参数
                nudAgentTemp.Value = (decimal)_appSettingsService.GetDoubleSetting("agent_temperature", 0.7);
                nudAgentMaxTokens.Value = _appSettingsService.GetIntSetting("agent_max_tokens", 8192);
                nudAgentTopP.Value = (decimal)_appSettingsService.GetDoubleSetting("agent_top_p", 0.9);
                
                System.Diagnostics.Debug.WriteLine("默认参数设置已加载");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载当前设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // 保存通用默认参数
                _appSettingsService.UpdateSetting("default_temperature", nudDefaultTemp.Value.ToString());
                _appSettingsService.UpdateSetting("default_max_tokens", nudDefaultMaxTokens.Value.ToString());
                _appSettingsService.UpdateSetting("default_top_p", nudDefaultTopP.Value.ToString());
                
                // 保存Chat模式参数
                _appSettingsService.UpdateSetting("chat_temperature", nudChatTemp.Value.ToString());
                _appSettingsService.UpdateSetting("chat_max_tokens", nudChatMaxTokens.Value.ToString());
                _appSettingsService.UpdateSetting("chat_top_p", nudChatTopP.Value.ToString());
                
                // 保存Agent模式参数
                _appSettingsService.UpdateSetting("agent_temperature", nudAgentTemp.Value.ToString());
                _appSettingsService.UpdateSetting("agent_max_tokens", nudAgentMaxTokens.Value.ToString());
                _appSettingsService.UpdateSetting("agent_top_p", nudAgentTopP.Value.ToString());
                
                System.Diagnostics.Debug.WriteLine("默认参数设置已保存到数据库");
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("确定要重置为系统默认值吗？", "确认重置", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
            if (result == DialogResult.Yes)
            {
                // 重置为系统默认值
                nudDefaultTemp.Value = 0.7m;
                nudDefaultMaxTokens.Value = 4000;
                nudDefaultTopP.Value = 0.9m;
                
                nudChatTemp.Value = 0.5m;
                nudChatMaxTokens.Value = 2048;
                nudChatTopP.Value = 0.8m;
                
                nudAgentTemp.Value = 0.7m;
                nudAgentMaxTokens.Value = 8192;
                nudAgentTopP.Value = 0.9m;
            }
        }
    }
} 