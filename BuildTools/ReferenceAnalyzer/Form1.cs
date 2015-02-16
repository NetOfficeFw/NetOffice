using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOBuildTools.ReferenceAnalyzer
{
    /// <summary>
    /// Main form in the application
    /// </summary>
    public partial class Form1 : Form
    {
        #region

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        #endregion

        #region Methods

        private void LogAction(string message)
        {
            if (!String.IsNullOrWhiteSpace(message))
                RichTextBoxLog.Text = message + Environment.NewLine + RichTextBoxLog.Text;
            this.Refresh();
            Application.DoEvents(); // not nice but okay for this simple one
        }

        #endregion

        #region Trigger

        private void ButttonChooseFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Xml Files(*.xml)|*.xml";
            if (DialogResult.OK == dialog.ShowDialog(this))
                TextBoxFile.Text = dialog.FileName;
        }

        private void ButtonStart_Click(object sender, EventArgs e)
        {
            try
            {
                if (String.IsNullOrWhiteSpace(TextBoxFile.Text))
                    return;
                RichTextBoxLog.Clear();
                XDocument document = Parser.ParseReference(LogAction);
                if (File.Exists(TextBoxFile.Text))
                    File.Delete(TextBoxFile.Text);
                document.Save(TextBoxFile.Text);
            }
            catch (Exception exception)
            {
                ExceptionDisplayer.ShowException(this, exception);
            }
        }

        #endregion
    }
}
