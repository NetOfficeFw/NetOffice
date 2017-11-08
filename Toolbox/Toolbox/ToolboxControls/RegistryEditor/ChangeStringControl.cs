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
    /// String Value Editor
    /// </summary>
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeStringDialogMessageTable.txt")]
    public partial class ChangeStringControl : UserControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ChangeStringControl()
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
        /// Value name
        /// </summary>
        public string EntryName
        {
            get
            {
                return textBoxName.Text;
            }
        }

        /// <summary>
        /// New value
        /// </summary>
        public string EntryValue
        {
            get
            {
                return textBoxValue.Text;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Set arguments to edit
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="value">value to edit</param>
        public void SetArguments(string name, string value)
        {
            textBoxName.Text = name;
            textBoxValue.Text = value;
        }

        /// <summary>
        /// Set focus to edit value
        /// </summary>
        public void SetFocus()
        {
            textBoxValue.Focus();
        }

        #endregion

        #region Trigger

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

        private void buttonOK_Click(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = DialogResult.OK;
                RaiseClose();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
