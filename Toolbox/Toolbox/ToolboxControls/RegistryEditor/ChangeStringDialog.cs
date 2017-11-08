using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    /// <summary>
    /// String edit host dialog
    /// </summary>
    partial class ChangeStringDialog : Form
    {
        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="value">value to edit</param>
        /// <param name="currentLanguageID">current user language id</param>
        public ChangeStringDialog(string name, string value, int currentLanguageID)
        {
            InitializeComponent();
            changeStringControl1.SetArguments(name, value);
            //Translation.Translator.AutoTranslateControls(changeStringControl1, "Registry Editor - ChangeString", "ToolboxControls.RegistryEditor.ChangeStringDialogMessageTable.txt", currentLanguageID);
            this.Text = changeStringControl1.Text;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Value name
        /// </summary>
        public string EntryName
        {
            get 
            {
                return changeStringControl1.EntryName;
            } 
        }

        /// <summary>
        /// New value
        /// </summary>
        public string EntryValue
        {
            get
            {
                return changeStringControl1.EntryValue;
            }
        }

        #endregion

        #region Trigger

        private void changeStringControl1_Close(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = changeStringControl1.DialogResult;
                this.Close();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void ChangeStringDialog_KeyDown(object sender, KeyEventArgs e)
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
