using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Show multi-line/rich text to the user
    /// </summary>
    public partial class RichTextDialog : ToolsDialog
    {
        #region Fields

        /// <summary>
        /// Save checkbox using explicitly because its unsafe to check for visibilty in result
        /// </summary>
        private bool _useCondition = false;

        private bool _skipOnUserAction;

        private bool _timeoutMode;

        private int _timeoutSeconds;

        private DateTime _startTime;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public RichTextDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">multi-line/rich text to display</param>
        /// <param name="timeoutSeconds">automatic close timeout</param>
        /// <param name="skipOnUserAction">skip timeout on user action</param>
        public RichTextDialog(string caption, string text, int timeoutSeconds, bool skipOnUserAction)
        {
            InitializeComponent();
            labelHeaderCaption.Text = caption;
            if (IsRichText(text))
                richTextBoxText.Rtf = text;
            else
                richTextBoxText.Text = text;
            _useCondition = false;
            checkBoxCondition.Visible = false;
            _startTime = DateTime.Now;

            _timeoutSeconds = timeoutSeconds;
            _skipOnUserAction = skipOnUserAction;

            if (_timeoutSeconds >= 1)
            {
                _timeoutSeconds = timeoutSeconds;
                _timeoutMode = true;
                _skipOnUserAction = skipOnUserAction;
                CloseTimer.Enabled = true;
                labelTimeLeft.Visible = true;
                CloseTimer.Tick += delegate
                {
                    int totalSeconds = Convert.ToInt32((DateTime.Now - _startTime).TotalSeconds);
                    labelTimeLeft.Text = String.Format("Close automatically in {0} second(s)", _timeoutSeconds - totalSeconds);
                    if (totalSeconds >= _timeoutSeconds)
                    {
                        CloseTimer.Enabled = false;
                        DoClose();
                    }
                };
            }
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">multi-line/rich text to display</param>
        public RichTextDialog(string caption, string text)
        {
            InitializeComponent();
            labelHeaderCaption.Text = caption;
            if (IsRichText(text))
                richTextBoxText.Rtf = text;
            else
                richTextBoxText.Text = text;
            _useCondition = false;
            checkBoxCondition.Visible = false;
            _startTime = DateTime.Now;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="caption">header caption on top</param>
        /// <param name="text">multi-line/rich text to display</param>
        /// <param name="checkText">additional condition text. checkbox is not shown if its null/empty</param>
        /// <param name="timeoutSeconds">automatic close timeout</param>
        /// <param name="skipOnUserAction">skip timeout on user action</param>
        public RichTextDialog(string caption, string text, string checkText, int timeoutSeconds, bool skipOnUserAction)
        {
            InitializeComponent();
            labelHeaderCaption.Text = caption;
            if(IsRichText(text))
                richTextBoxText.Rtf = text;
            else
                richTextBoxText.Text = text;

            _startTime = DateTime.Now;
            if (!String.IsNullOrEmpty(checkText))
            {
                _useCondition = true;
                checkBoxCondition.Text = checkText;
            }
            else
            {
                _useCondition = false;
                checkBoxCondition.Visible = false;
            }

            _timeoutSeconds = timeoutSeconds;
            _skipOnUserAction = skipOnUserAction;

            if (_timeoutSeconds >= 1)
            {
                _timeoutSeconds = timeoutSeconds;
                _timeoutMode = true;
                _skipOnUserAction = skipOnUserAction;
                CloseTimer.Enabled = true;
                labelTimeLeft.Visible = true;
                CloseTimer.Tick += delegate
                {
                    int totalSeconds = Convert.ToInt32((DateTime.Now - _startTime).TotalSeconds);
                    labelTimeLeft.Text = String.Format("Close automatically in {0} second(s)", _timeoutSeconds - totalSeconds);
                    if (totalSeconds >= _timeoutSeconds)
                    {
                        CloseTimer.Enabled = false;
                        DoClose();
                    }
                };
            }
        }

        #endregion

        #region Methods

        private static bool IsRichText(string text)
        {
            if (String.IsNullOrEmpty(text))
                return false;
            if (text.TrimStart().StartsWith("{\rtf1", StringComparison.Ordinal))
                return true;
            else
                return false;      
        }

        private void DoClose()
        {
            CloseTimer.Enabled = false;
            if (_useCondition)
                DialogResult = checkBoxCondition.Checked ? DialogResult.OK : DialogResult.Cancel;
            else
                DialogResult = DialogResult.OK;

            this.Close();
        }

        #endregion

        #region Trigger

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                DoClose();
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);                
            }
        }

        private void This_Click(object sender, EventArgs e)
        {
            if (_skipOnUserAction && _timeoutMode)
            {
                labelTimeLeft.Visible = false;
                CloseTimer.Enabled = false;
                _timeoutMode = false;
            }         
        }

        private void richTextBoxText_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(e.LinkText);
            }
            catch
            {
                ;
            }
        }

        #endregion
    }
}
