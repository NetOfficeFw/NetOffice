using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeNameDialogMessageTable.txt")]
    partial class ChangeNameDialog : Form
    {
        public ChangeNameDialog(string name, int currentLanguageID)
        {
            InitializeComponent();
            changeNameControl1.SetArguments(name);
            Translation.Translator.AutoTranslateControls(changeNameControl1, "Registry Editor - ChangeName", "ToolboxControls.RegistryEditor.ChangeNameDialogMessageTable.txt", currentLanguageID);
            this.Text = changeNameControl1.Text;
            changeNameControl1.SetFocus();
        }

        public string EntryNewName
        {
            get 
            {
                return changeNameControl1.EntryNewName.Trim();
            }
        }

        #region Trigger

        private void changeNameControl1_Close(object sender, EventArgs e)
        {
            this.DialogResult = changeNameControl1.DialogResult;
            this.Close();
        }

        private void ChangeNameDialog_KeyDown(object sender, KeyEventArgs e)
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

        #endregion
    }
}
