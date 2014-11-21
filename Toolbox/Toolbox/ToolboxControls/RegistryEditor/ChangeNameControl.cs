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
    [RessourceTable("ToolboxControls.RegistryEditor.ChangeNameDialogMessageTable.txt")]
    public partial class ChangeNameControl : UserControl, ILocalizationDesign
    {
        #region Ctor

        public ChangeNameControl()
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

        public string EntryNewName
        {
            get
            {
                return textBoxValue.Text.Trim();
            }
        }

        #endregion

        #region Methods

        public void SetArguments(string name)
        {
            textBoxName.Text = name;
            textBoxValue.Text = name;
        }

        public void SetFocus()
        {
            textBoxValue.Focus();
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
            if (textBoxValue.Text.Trim() != textBoxName.Text)
                this.DialogResult = DialogResult.OK;
            RaiseClose();
        }

        private void buttonAbort_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            RaiseClose();
        }

        private void textBoxValue_TextChanged(object sender, EventArgs e)
        {
            buttonOK.Enabled = !String.IsNullOrWhiteSpace(textBoxValue.Text);
        }

        #endregion
    }
}
