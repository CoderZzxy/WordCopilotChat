using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WordCopilot.utils;

namespace WordCopilotChat
{

    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        Common common = new Common();  // 实例化类
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            common.ShowCustomTask();
        }

        private void btnModel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var modelForm = new ModelListForm();
                modelForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开模型列表失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnPrompt_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var promptForm = new PromptSettingsForm();
                promptForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开提示词设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDoc_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var docForm = new DocumentSettingsForm();
                docForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开提文档管理失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void checkBoxLog_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 仅当勾选时才开启日志存储；默认未勾选（关闭），以节省磁盘空间
                bool enabled = this.checkBoxLog.Checked;
                OpenAIUtils.EnableLogging = enabled;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"切换日志存储失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
    
}
