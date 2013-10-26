using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.RegistryEditor
{
    partial class ChangeNameDialog : Form
    {
        public ChangeNameDialog(string name, int currentLanguageID)
        {
            InitializeComponent();
            textBoxName.Text = name;
            textBoxValue.Text = name;
            Translator.TranslateControls(this, "RegistryEditor.ChangeNameDialogMessageTable.txt", currentLanguageID);
        }

        public string EntryNewName
        {
            get 
            {
                return textBoxValue.Text.Trim();
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (textBoxValue.Text.Trim() != textBoxName.Text)
                this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
