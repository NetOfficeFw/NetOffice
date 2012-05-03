using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.ProjectWizard
{
    partial class ErrorDialog : Form
    {
        public ErrorDialog(Exception exception, TargetLanguage language)
        {
            InitializeComponent();
            Translator.TranslateControls(this, "Dialogs.ErrorDialog.txt", language);

            string errorLog = "";
            while (null != exception)
            {
                errorLog += exception.Message + Environment.NewLine;
                exception = exception.InnerException;
            }

            textBoxErrorLog.Text = errorLog;

        }

        private void okayButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void copyButton_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBoxErrorLog.Text);
        }

        private void linkLabelNetOfficeDiscussion_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://netoffice.codeplex.com/discussions");
            }
            catch
            {
                MessageBox.Show("Browser konnte nicht gestartet werden", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
