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
            textBoxName.Text = name;
            DynamicByteProvider provider = new DynamicByteProvider(value);
            hexBox.ByteProvider = provider;
            Translation.Translator.TranslateControls(this, "ToolboxControls.RegistryEditor.ChangeBinaryDialogMessageTable.txt", currentLanguageID);
        }
        
        #endregion

        #region Properties

        public Byte[] Bytes
        {
            get 
            {
                DynamicByteProvider provider = hexBox.ByteProvider as DynamicByteProvider;
                return provider.Bytes.ToArray();
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

        private void ChangeBinaryDialog_Resize(object sender, EventArgs e)
        {
             hexBox.BytesPerLine = this.Width / 56;
        }

        #endregion

        #region Helper

        private string ByteArrayToString(byte[] byteArray)
        {
            string result = "";
            foreach (byte value in byteArray)
            {
               char output = Convert.ToChar(value);
               result += output;
            }
            return result;
        }

        private static byte[] StringToByteArray(string str)
        {
            if (null == str)
                return null;
            System.Text.UnicodeEncoding enc = new System.Text.UnicodeEncoding();
            return enc.GetBytes(str);
        }

        #endregion
    }
}
