using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeUI
{
    public delegate void SelectOfficeEventHandler(string officeAppName);

    [RessourceTable("ToolboxControls.OfficeUI.SelectOfficeAppControlStrings.txt")]
    public partial class SelectOfficeAppControl : UserControl, ILocalizationDesign
    {
        SelectOfficeEventHandler _eventHandler;
        int _currentLanguageID;
        
        public SelectOfficeAppControl( )
        {
            InitializeComponent();
        }

        public SelectOfficeAppControl(int currentLanguageID, SelectOfficeEventHandler handler)
        {
            InitializeComponent();
            _eventHandler = handler;
            _currentLanguageID = currentLanguageID;
            Translation.Translator.TranslateControls(this, "ToolboxControls.OfficeUI.SelectOfficeAppControlStrings.txt", _currentLanguageID);
        }

        public string SelectedApplication
        {
            get 
            {
                return listView1.SelectedItems[0].Text;
            }
        }

        public DialogResult Result { get; set; }
        
        private void buttonClose2_Click(object sender, EventArgs e)
        {
            Result = DialogResult.Cancel;
            this.Hide();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Result = DialogResult.Cancel;
            this.Hide();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonSelect.Enabled  = (listView1.SelectedIndices.Count > 0);            
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if ((listView1.SelectedIndices.Count > 0))
                buttonSelect_Click(this, new EventArgs());
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            try
            {
                Result = DialogResult.OK;
                this.Hide();
                _eventHandler(SelectedApplication);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, _currentLanguageID);
            }
        }

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
            get { throw new NotImplementedException(); }
        }

        #endregion

    }
}
