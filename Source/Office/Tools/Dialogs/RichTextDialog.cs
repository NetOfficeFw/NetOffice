using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
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
        /// <param name="checkText">additional condition text. checkbox is not shown if its null/empty</param>
        public RichTextDialog(string caption, string text, string checkText)
        {
            InitializeComponent();
            labelHeaderCaption.Text = caption;
            richTextBoxText.Text = text;
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
        }

        #endregion

        #region Trigger

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {                
                if (_useCondition)
                    DialogResult = checkBoxCondition.Checked ? DialogResult.OK : DialogResult.Cancel;
                else
                    DialogResult = DialogResult.OK;

                this.Close();
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);                
            }
        }

        #endregion
    }
}
