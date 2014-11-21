using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.Controls.Hex;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    partial class ChangeBinaryDialog : Form
    {
        #region Construction

        public ChangeBinaryDialog(string name, byte[] value, int currentLanguageID)
        {
            InitializeComponent();
            changeBinaryControl1.SetArguments(name, value);
            Translation.Translator.AutoTranslateControls(changeBinaryControl1, "Registry Editor - ChangeBinary", "ToolboxControls.RegistryEditor.ChangeBinaryDialogMessageTable.txt", currentLanguageID);
            this.Text = changeBinaryControl1.Text;
        }
        
        #endregion

        #region Properties

        public Byte[] Bytes
        {
            get 
            {
                return changeBinaryControl1.Bytes;
            }
        }

        #endregion


        #region Trigger

        private void ChangeBinaryDialog_KeyDown(object sender, KeyEventArgs e)
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

        private void changeBinaryControl1_Close(object sender, EventArgs e)
        {
            this.DialogResult = changeBinaryControl1.DialogResult;
            this.Close();
        }
        
        #endregion
    }
}
