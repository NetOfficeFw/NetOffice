using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Translation
{
    public partial class LanguageSummaryControl : UserControl
    {
        private ToolLanguage _selectedLanguage;
        private bool _initialize;
        private bool _firstChangePassed;

        public LanguageSummaryControl()
        {
            InitializeComponent();
        }

        internal ToolLanguage SelectedLanguage
        {
            get
            {
                return _selectedLanguage;
            }
            set 
            {
                _selectedLanguage = value;
                ShowLanguage();
            }
        }

        private void ShowLanguage()
        {
            _initialize = true;
            if (null != _selectedLanguage)
            {
                textBoxNameGlobal.DataBindings.Add("Text", _selectedLanguage, "NameGlobal", true, DataSourceUpdateMode.OnPropertyChanged);
                textBoxNameLocal.DataBindings.Add("Text", _selectedLanguage, "Name", true, DataSourceUpdateMode.OnPropertyChanged);
                textBoxLanguageID.DataBindings.Add("Text", _selectedLanguage, "LCID", true, DataSourceUpdateMode.OnPropertyChanged);
                textBoxAuthorName.DataBindings.Add("Text", _selectedLanguage, "Author", true, DataSourceUpdateMode.OnPropertyChanged);
                textBoxAuthorSite.DataBindings.Add("Text", _selectedLanguage, "AuthorSite", true, DataSourceUpdateMode.OnPropertyChanged);
                textBoxAuthorMail.DataBindings.Add("Text", _selectedLanguage, "AuthorMail", true, DataSourceUpdateMode.OnPropertyChanged);
                _selectedLanguage.PropertyChanged += new PropertyChangedEventHandler(Item_PropertyChanged);            
            }
            else
            {
                textBoxNameGlobal.Text = String.Empty;
                textBoxNameLocal.Text = String.Empty;
                textBoxLanguageID.Text = "0";
                textBoxAuthorName.Text = String.Empty;
                textBoxAuthorSite.Text = String.Empty;
                textBoxAuthorMail.Text = String.Empty;
            }
            _initialize = false;
        }
        
        private string GetLanguageLCID(string languageName)
        {
            string countriesContent = Ressources.RessourceUtils.ReadString("Translation.Countries.txt");
            if (null == countriesContent)
                return null;
            int languageIndex = -1;
            string[] array = countriesContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < array.Length; i++)
			{
                if (array[i].Trim().Equals(languageName, StringComparison.InvariantCultureIgnoreCase))
                {
                    languageIndex = i;
                    break;
                }
            }
            
            if (languageIndex < 0)
                return null;

            string lcidContent = Ressources.RessourceUtils.ReadString("Translation.LCIDs.txt");
            if (null == lcidContent)
                return null;
            string[] array2 = lcidContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            if (array.Length <= array2.Length)
                return null;
            return array2[languageIndex-1];
        }


        private void Item_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            _selectedLanguage.IsDirty = true;
        }

        private decimal ToDecimal(string value)
        { 
            decimal d = 0;
            decimal.TryParse(value, out d);
            if (d < 0 || d > 20000)
                return 0;
            else
                return d;
        }

        private void linkLabelLCID_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                string link = linkLabelLCID.Tag as string;
                System.Diagnostics.Process.Start(link);
            }
            catch
            {
                ;
            }
        }

        private void textBoxNameGlobal_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (_initialize)
                    return;
                if (!_firstChangePassed)
                {
                    _firstChangePassed = true;
                    return;
                }

                string lcid = GetLanguageLCID(textBoxNameGlobal.Text.Trim());
                if (null != lcid)
                {
                    textBoxLanguageID.Text = lcid;
                    textBoxNameLocal.Text = textBoxNameGlobal.Text;
                }
            }
            catch 
            {
                ;
            }
        }
    }
}
