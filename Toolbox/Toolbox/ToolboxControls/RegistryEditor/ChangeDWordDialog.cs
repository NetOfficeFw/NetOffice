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
            textBoxName.Text = name;
            textBoxValue.Text = InitialConvertValue(value);
            Translation.Translator.TranslateControls(this, "ToolboxControls.RegistryEditor.ChangeDWordDialogMessageTable.txt", currentLanguageID);
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

        public object EntryValue
        {
            get
            {
                return FinalConvertValue(textBoxValue.Text);
            }
        }

        #endregion

        #region Trigger

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBoxValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ("1234567890,\b".IndexOf(e.KeyChar.ToString()) < 0)
                e.Handled = true;
        }

        private void radioButtonHex_CheckedChanged(object sender, EventArgs e)
        {
            object value = ConvertValue(textBoxValue.Text);
            textBoxValue.Text = value.ToString() ;
        }

        #endregion

        #region Methods

        private object ConvertValue(string value)
        {
           
            if (radioButtonHex.Checked == true)
                return ConvertDecimalStringToHexValue(value);
            else
                return ConvertHexStringToDecimal(value);
        }

        private object ConvertDecimalStringToHexValue(string value)
        {
            int objectValue = Convert.ToInt32(value);
            return string.Format("{0:x}", objectValue);
        }

        private object ConvertHexStringToDecimal(string value)
        {
            return Int32.Parse(value, System.Globalization.NumberStyles.HexNumber);
        }

        private string InitialConvertValue(object value)
        {
            if (radioButtonHex.Checked == true)
            {
                int val = Convert.ToInt32(value);
                return string.Format("{0:x}", value);
            }
            else
                return value.ToString();
        }

        private string FinalConvertValue(string value)
        {
            if (radioButtonHex.Checked == true)
            {
                int decValue = Int32.Parse(value, System.Globalization.NumberStyles.HexNumber);
                return decValue.ToString();
            }
            else
            { 
                return value.ToString();
            }
        }

        #endregion
    }
}
