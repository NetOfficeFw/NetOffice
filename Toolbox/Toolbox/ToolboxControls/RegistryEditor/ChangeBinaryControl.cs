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
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeBinaryDialogMessageTable.txt")]
    public partial class ChangeBinaryControl : UserControl, ILocalizationDesign
    {
        #region Ctor

        public ChangeBinaryControl()
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

        public Byte[] Bytes
        {
            get
            {
                DynamicByteProvider provider = hexBox.ByteProvider as DynamicByteProvider;
                return provider.Bytes.ToArray();
            }
        }

        #endregion

        #region Methods

        public void SetArguments(string name, byte[] value)
        {
            DynamicByteProvider provider = new DynamicByteProvider(value);
            hexBox.ByteProvider = provider;
            textBoxName.Text = name;
            hexBox.ByteProvider = provider;
        }

        public void SetFocus()
        {
            hexBox.Focus();
        }

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

        #region Trigger

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            RaiseClose();
        }

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            RaiseClose();
        }

        private void ChangeBinaryControl_Resize(object sender, EventArgs e)
        {
            hexBox.BytesPerLine = this.Width / 56;
        }

        #endregion
    }
}
