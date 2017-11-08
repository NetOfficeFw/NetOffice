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
    /// Name edit host dialog
    /// </summary>
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeNameDialogMessageTable.txt")]
    partial class ChangeNameDialog : Form
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">name to edit</param>
        /// <param name="currentLanguageID">current user language id</param>
        public ChangeNameDialog(string name, int currentLanguageID)
        {
            InitializeComponent();
            changeNameControl1.SetArguments(name);
            //Translation.Translator.AutoTranslateControls(changeNameControl1, "Registry Editor - ChangeName", "ToolboxControls.RegistryEditor.ChangeNameDialogMessageTable.txt", currentLanguageID);
            this.Text = changeNameControl1.Text;
            changeNameControl1.SetFocus();
        }

        #endregion

        #region Properties

        /// <summary>
        /// New name
        /// </summary>
        public string EntryNewName
        {
            get 
            {
                return changeNameControl1.EntryNewName.Trim();
            }
        }

        #endregion

        #region Trigger

        private void changeNameControl1_Close(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = changeNameControl1.DialogResult;
                this.Close();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void ChangeNameDialog_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Escape)
                {
                    this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                    this.Close();
                }
                else if (e.KeyCode == Keys.Return && false == String.IsNullOrWhiteSpace(EntryNewName))
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
