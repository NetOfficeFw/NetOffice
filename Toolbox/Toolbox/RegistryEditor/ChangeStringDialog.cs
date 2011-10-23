using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.RegistryEditor
{
    partial class ChangeStringDialog : Form
    {
        #region Construction

        public ChangeStringDialog(string name, string value, int currentLanguageID)
        {
            InitializeComponent();
            textBoxName.Text = name;
            textBoxValue.Text = value;
            Translator.TranslateControls(this, "RegistryEditor.ChangeStringDialogMessageTable.txt", currentLanguageID);
            textBoxValue.Focus();
        }

        #endregion

        #region Properties

        public string EntryName
        {
            get 
            {
                return textBoxName.Text;
            } 
        }

        public string EntryValue
        {
            get
            {
                return textBoxValue.Text;
            }
        }

        #endregion

        #region Trigger

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion
    }
}
