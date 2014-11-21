using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeDWordDialogMessageTable.txt")]
    public partial class ChangeDWORDControl : UserControl, ILocalizationDesign
    {
        #region Ctor

        public ChangeDWORDControl()
        {
            InitializeComponent();
        }

        #endregion

        #region Events

        public event EventHandler Close;

        private void RaiseClose()
        {
            if (null != Close)
                Close(this, EventArgs.Empty);
        }

        #endregion

        #region Properties

        public DialogResult DialogResult { get; private set; }

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

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {

        }

        public void Localize(Translation.ItemCollection strings)
        {
            Translation.Translator.TranslateControls(this, strings);
        }

        public void Localize(string name, string text)
        {
            Translation.Translator.TranslateControl(this, name, text);
        }

        public string GetCurrentText(string name)
        {
            return Translation.Translator.TryGetControlText(this, name);
        }

        public IContainer Components
        {
            get { return components; }
        }

        public string NameLocalization
        {
            get
            {
                return null;
            }
        }

        public IEnumerable<ILocalizationChildInfo> Childs
        {
            get
            {
                return new ILocalizationChildInfo[0];
            }
        }

        #endregion

        #region Methods

        public void SetArguments(string name, object value)
        {
            textBoxName.Text = name;
            textBoxValue.Text = InitialConvertValue(value);
        }

        public void SetFocus()
        {
            textBoxValue.Focus();
        }

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

        #region Trigger

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            RaiseClose();
        }

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            RaiseClose();
        }

        private void textBoxValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ("1234567890,\b".IndexOf(e.KeyChar.ToString()) < 0)
                e.Handled = true;
        }

        private void radioButtonHex_CheckedChanged(object sender, EventArgs e)
        {
            object value = ConvertValue(textBoxValue.Text);
            textBoxValue.Text = value.ToString();
        }

        #endregion
    }
}
