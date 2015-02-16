using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace NOBuildTools.SearchAndReplace
{
    /// <summary>
    /// Main form in the application
    /// </summary>
    public partial class Form1 : Form
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            TextBoxFolder.Text = Application.StartupPath;
        }

        #endregion

        #region Methods

        private void LogAction(string message)
        {
            if(!String.IsNullOrWhiteSpace(message))
                RichTextBoxLog.Text = message + Environment.NewLine + RichTextBoxLog.Text;
            this.Refresh();
        }

        #endregion

        #region Triger

        private void ButtonChooseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.Cancel)
                return;
            TextBoxFolder.Text = dlg.SelectedPath;
        }

        private void ButtonStart_Click(object sender, EventArgs e)
        {
            try
            {
                RichTextBoxLog.Clear();
                SearchAndReplaceManager.SearchAndReplace(TextBoxFolder.Text, TextBoxFilter.Text, TextBoxSearch.Text, TextBoxReplace.Text, LogAction);
            }
            catch (Exception exception)
            {
                ExceptionDisplayer.ShowException(this, exception);
            }
        }

        private void ButtonSaveConfig_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.InitialDirectory = Application.StartupPath;
                dlg.Filter = "Xml Files(*.xml)|*.xml";
                if (dlg.ShowDialog(this) == DialogResult.Cancel)
                    return;

                ConfigManager.SaveConfigurationToXMLFile(dlg.FileName, TextBoxFolder.Text, TextBoxFilter.Text, TextBoxSearch.Text, TextBoxReplace.Text);
            }
            catch (Exception exception)
            {
                ExceptionDisplayer.ShowException(this, exception);
            }
        }

        private void ButtonLoadConfig_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.InitialDirectory = Application.StartupPath;
                dlg.Filter = "Xml Files(*.xml)|*.xml";
                if (dlg.ShowDialog(this) == DialogResult.Cancel)
                    return;

                string targetFolder = string.Empty;
                string filter = string.Empty;
                string search = string.Empty;
                string replace = string.Empty;

                ConfigManager.LoadConfigurationFromConfigFile(dlg.FileName, ref targetFolder, ref filter, ref search, ref replace);

                TextBoxFolder.Text = targetFolder;
                TextBoxFilter.Text = filter;
                TextBoxSearch.Text = search;
                TextBoxReplace.Text = replace;
            }
            catch (Exception exception)
            {
                ExceptionDisplayer.ShowException(this, exception);
            }
        }

        #endregion
    }
}
