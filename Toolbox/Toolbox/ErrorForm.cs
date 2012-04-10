using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    public enum ErrorCategory
    { 
        NonCritical = 0,
        Critical = 1,
        Penalty = 2
    }

    partial class ErrorForm : Form
    {
        #region Fields

        ErrorCategory _category;
        bool _isExpanded;

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
            Translator.TranslateControls(this, "ErrorFormMessageTable.txt", currentLanguageID);
            this.Height = buttonOK.Top + buttonOK.Height + 40;
        }

        public ErrorForm(Exception exception, ErrorCategory category, int currentLanguageID)
        {
            InitializeComponent();
            _category = category;
            if (ErrorCategory.Critical == category)
                labelExitMessage.Visible = true;
            DisplayException(exception);
            currentLanguageID = ValidateLanguageID(currentLanguageID);
            Translator.TranslateControls(this, "ErrorFormMessageTable.txt", currentLanguageID);
            this.Height = buttonOK.Top + buttonOK.Height + 40;
        }

        #endregion

        #region Trigger

        private void buttonDetails_Click(object sender, EventArgs e)
        {
            if (_isExpanded)
            {
                this.Height = buttonOK.Top + buttonOK.Height + 40;
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
                listViewTrace.Anchor = AnchorStyles.None;
                listViewTrace.Left = 26;
                listViewTrace.Top = 130;
                listViewTrace.Width = 374;
                listViewTrace.Height = 164;
            }
            else
            { 
                this.Height = 360;
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                listViewTrace.Left = 26;
                listViewTrace.Top = 130;
                listViewTrace.Width = (buttonOK.Left + buttonOK.Width) - listViewTrace.Left;
                listViewTrace.Height = 164;
                listViewTrace.Anchor = AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
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

        #endregion

        #region Methods

        private int ValidateLanguageID(int currentLanguageID)
        {
            switch (currentLanguageID)
            {
                case 0:
                    currentLanguageID = 1033;
                    break;
                case 1:
                    currentLanguageID = 1031;
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
                exception = exception.InnerException;
                i++;
            }
        }

        #endregion
    }
}
