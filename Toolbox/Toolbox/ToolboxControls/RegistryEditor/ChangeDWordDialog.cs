using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.Utils.Registry;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    /// <summary>
    /// DWORD edit host dialog
    /// </summary>
    partial class ChangeDWordDialog : Form
    {
        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="value">value to edit</param>
        /// <param name="currentLanguageID">current user language id</param>
        public ChangeDWordDialog(string name, object value, int currentLanguageID)
        {
            InitializeComponent();
            changeDWORDControl1.SetArguments(name, value);
            Translation.Translator.AutoTranslateControls(changeDWORDControl1, "Registry Editor - ChangeDWORD", "ToolboxControls.RegistryEditor.ChangeDWordDialogMessageTable.txt", currentLanguageID);
            this.Text = changeDWORDControl1.Text;
            changeDWORDControl1.SetFocus();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Name of the value
        /// </summary>
        public string EntryName
        {
            get
            {
                return changeDWORDControl1.EntryName;
            }
        }

        /// <summary>
        /// Value to edit
        /// </summary>
        public object EntryValue
        {
            get
            {
                return changeDWORDControl1.EntryValue;
            }
        }

        #endregion

        #region Trigger

        private void changeDWORDControl1_Close(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = changeDWORDControl1.DialogResult;
                this.Close();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void ChangeDWORDDialog_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Escape)
                {
                    this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                    this.Close();
                }
                else if (e.KeyCode == Keys.Return)
                {
                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
