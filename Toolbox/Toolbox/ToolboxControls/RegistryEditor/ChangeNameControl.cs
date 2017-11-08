using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    /// <summary>
    /// Name Value Editor
    /// </summary>
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeNameDialogMessageTable.txt")]
    public partial class ChangeNameControl : UserControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ChangeNameControl()
        {
            InitializeComponent();
        }

        #endregion

        #region Events

        /// <summary>
        /// User want close the dialog
        /// </summary>
        public event EventHandler Close;

        private void RaiseClose()
        {
            if (null != Close)
                Close(this, EventArgs.Empty);
        }

        #endregion

        #region Properties

        /// <summary>
        /// User want proceed edit or abort
        /// </summary>
        public DialogResult DialogResult { get; private set; }

        /// <summary>
        /// New name
        /// </summary>
        public string EntryNewName
        {
            get
            {
                return textBoxValue.Text.Trim();
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Set name to edit
        /// </summary>
        /// <param name="name">name to edit</param>
        public void SetArguments(string name)
        {
            textBoxName.Text = name;
            textBoxValue.Text = name;
        }

        /// <summary>
        /// Set focus to name edit
        /// </summary>
        public void SetFocus()
        {
            textBoxValue.Focus();
        }

        #endregion

        #region Trigger

        private void buttonOK_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBoxValue.Text.Trim() != textBoxName.Text)
                    this.DialogResult = DialogResult.OK;
                RaiseClose();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = DialogResult.Cancel;
                RaiseClose();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void textBoxValue_TextChanged(object sender, EventArgs e)
        {
            try
            {
                buttonOK.Enabled = !String.IsNullOrWhiteSpace(textBoxValue.Text);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
