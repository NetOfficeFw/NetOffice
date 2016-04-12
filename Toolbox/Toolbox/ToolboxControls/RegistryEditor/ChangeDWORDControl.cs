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
    /// <summary>
    /// DWORD Value Editor
    /// </summary>
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeDWordDialogMessageTable.txt")]
    public partial class ChangeDWORDControl : UserControl, ILocalizationDesign
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ChangeDWORDControl()
        {
            InitializeComponent();
        }

        #endregion

        #region Events

        /// <summary>
        /// User want close the dialog
        /// </summary>
        public event EventHandler Close;

        private void RaiseClose()
        {
            if (null != Close)
                Close(this, EventArgs.Empty);
        }

        #endregion

        #region Properties

        /// <summary>
        /// User want proceed or abort edit
        /// </summary>
        public DialogResult DialogResult { get; private set; }

        /// <summary>
        /// Name of the value
        /// </summary>
        public string EntryName
        {
            get
            {
                return textBoxName.Text;
            }
        }

        /// <summary>
        /// Value to edit
        /// </summary>
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

        /// <summary>
        /// Set edit arguments
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="value">value to edit</param>
        public void SetArguments(string name, object value)
        {
            textBoxName.Text = name;
            textBoxValue.Text = InitialConvertValue(value);
        }

        /// <summary>
        /// Set focus to the edit value
        /// </summary>
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
            try
            {
                this.DialogResult = DialogResult.OK;
                RaiseClose();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            try
            {
                RaiseClose();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void textBoxValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if ("1234567890,\b".IndexOf(e.KeyChar.ToString()) < 0)
                    e.Handled = true;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void radioButtonHex_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                object value = ConvertValue(textBoxValue.Text);
                textBoxValue.Text = value.ToString();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
