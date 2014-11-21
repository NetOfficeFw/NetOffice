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
    partial class ChangeDWordDialog : Form
    {
        #region Construction

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

        public string EntryName
        {
            get
            {
                return changeDWORDControl1.EntryName;
            }
        }

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
            this.DialogResult = changeDWORDControl1.DialogResult;
            this.Close();
        }

        private void ChangeDWORDDialog_KeyDown(object sender, KeyEventArgs e)
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
