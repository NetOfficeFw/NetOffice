using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    partial class ChangeStringDialog : Form
    {
        #region Construction

        public ChangeStringDialog(string name, string value, int currentLanguageID)
        {
            InitializeComponent();
            changeStringControl1.SetArguments(name, value);
            Translation.Translator.AutoTranslateControls(changeStringControl1, "Registry Editor - ChangeString", "ToolboxControls.RegistryEditor.ChangeStringDialogMessageTable.txt", currentLanguageID);
            this.Text = changeStringControl1.Text;
        }

        #endregion

        #region Properties

        public string EntryName
        {
            get 
            {
                return changeStringControl1.EntryName;
            } 
        }

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
            this.DialogResult = changeStringControl1.DialogResult;
            this.Close();
        }

        private void ChangeStringDialog_KeyDown(object sender, KeyEventArgs e)
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

        #endregion
    }
}
