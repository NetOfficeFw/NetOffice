using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Forms
{
    partial class ErrorForm : Form
    {
        #region Fields

        private ErrorCategory _category;
        private bool _isExpanded;

        #endregion

        #region Construction

        public ErrorForm(Exception exception, string message, ErrorCategory category, int currentLanguageID)
        {
            InitializeComponent();
            _category = category;
            labelErrorMessage.Text = message;
            labelErrorMessage.Visible = true;
            if (ErrorCategory.Critical == category)
                labelExitMessage.Visible = true;
            DisplayException(exception);
            currentLanguageID = ValidateLanguageID(currentLanguageID);
            Translation.Translator.TranslateControls(this, "Ressources.ErrorFormStrings.txt", currentLanguageID);
            this.Height = pictureBoxSplitter1.Top + (ClientRectangle.Height - DisplayRectangle.Height);
        }

        public ErrorForm(Exception exception, ErrorCategory category, int currentLanguageID)
        {
            InitializeComponent();
            _category = category;
            if (ErrorCategory.Critical == category)
                labelExitMessage.Visible = true;
            DisplayException(exception);
            currentLanguageID = ValidateLanguageID(currentLanguageID);
            Translation.Translator.TranslateControls(this, "Ressources.ErrorFormStrings.txt", currentLanguageID);
            this.Height = pictureBoxSplitter1.Top + (ClientRectangle.Height - DisplayRectangle.Height);
        }

        #endregion

        #region Methods

        public static void ShowError(Exception exception, ErrorCategory category, int currentLanguageID)
        {
            ErrorForm form = new ErrorForm(exception, category, currentLanguageID);
            if (null != MainForm.Singleton && MainForm.Singleton.Visible)
                form.ShowDialog(MainForm.Singleton);
            else
            {
                form.StartPosition = FormStartPosition.CenterScreen;
                form.ShowDialog();
            }
        }

        public static void ShowError(IWin32Window parent, Exception exception, ErrorCategory category, int currentLanguageID)
        {
            ErrorForm form = new ErrorForm(exception, category, currentLanguageID);
            form.ShowDialog(parent);
        }

        public static void ShowError(IWin32Window parent, Exception exception, ErrorCategory category)
        {
            ErrorForm form = new ErrorForm(exception, category, 1033 );
            form.ShowDialog(parent);
        }

        public static void ShowError(IWin32Window parent, Exception exception)
        {
            ErrorForm form = new ErrorForm(exception, Forms.ErrorCategory.NonCritical, 1033);
            form.ShowDialog(parent);
        }

        private int ValidateLanguageID(int currentLanguageID)
        {
            switch (currentLanguageID)
            {
                case 1:
                    currentLanguageID = 1031;
                    break;
                default:
                    currentLanguageID = 1033;
                    break;

            }

            return currentLanguageID;
        }

        private void DisplayException(Exception exception)
        {
            int i = 1;
            while (exception != null)
            {
                ListViewItem viewItem = listViewTrace.Items.Add(i.ToString());
                viewItem.SubItems.Add(exception.Message);
                viewItem.SubItems.Add(exception.GetType().Name.ToString());
                if (null != exception.TargetSite)
                    viewItem.SubItems.Add(exception.TargetSite.ToString());
                else
                    viewItem.SubItems.Add("");
                viewItem.Tag = exception;
                exception = exception.InnerException;
                i++;
            }
        }

        #endregion

        #region Trigger

        private void buttonDetails_Click(object sender, EventArgs e)
        {
            if (_isExpanded)
            {
                this.Height = pictureBoxSplitter1.Top + (ClientRectangle.Height - DisplayRectangle.Height);
            }
            else
            {
                this.Height = pictureBoxSplitter2.Top + pictureBoxSplitter2.Height + (ClientRectangle.Height - DisplayRectangle.Height);
            }
            _isExpanded = !_isExpanded;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Close();
            if (ErrorCategory.Critical == _category)
                Application.Exit();
        }

        private void buttonCopyToClipboard_Click(object sender, EventArgs e)
        {
            string clipboardContent = "";

            foreach (ListViewItem item in listViewTrace.Items)
                clipboardContent += item.SubItems[0].Text + " | " + item.SubItems[1].Text + " | " + item.SubItems[2].Text + " | " + item.SubItems[3].Text + Environment.NewLine;

            Clipboard.SetData(DataFormats.Text, clipboardContent); 
        }

        private void linkLabelDiscussionBoard_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start((sender as LinkLabel).Tag as string);
            }
            catch
            {
                ;
            }
        }

        private void listViewTrace_DoubleClick(object sender, EventArgs e)
        {
            if (listViewTrace.SelectedItems.Count > 0)
            {
                Exception exception = listViewTrace.SelectedItems[0].Tag as Exception;
                if (null != exception)
                {
                    string details = String.Format("{0}{2}{2}{1}", exception.Message, exception, Environment.NewLine);
                    MessageBox.Show(this, details, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        #endregion
    }

    /// <summary>
    /// define error categories
    /// </summary>
    public enum ErrorCategory
    {
        /// <summary>
        /// the error is non critical
        /// </summary>
        NonCritical = 0,

        /// <summary>
        /// the error is an critical/unexpected error
        /// </summary>
        Critical = 1,

        /// <summary>
        /// the error is a sudden death error. the program has to terminate immediately
        /// </summary>
        Penalty = 2
    }
}
