using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    //step  2
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.EnvironmentControl.txt")]
    public partial class EnvironmentControl : UserControl, IWizardControl, ILocalizationDesign
    {
        private XmlDocument _settings;

        public EnvironmentControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
            comboBoxNetRuntime.SelectedIndex = 3;
        }

        public event ReadyStateChangedHandler ReadyStateChanged;
        
        public new void KeyDown(KeyEventArgs e)
        { 
        }

        public bool IsReadyForNextStep
        {
            get { return true; }
        }

        public string Caption
        {
            get
            {
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
                    return "Umgebung";
                else
                    return "Environment";
            }
        }

        public string Description
        {
            get
            {
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
                    return "Der Assistent kann die Implementierung für Sie vorbereiten.";
                else
                    return "The assistent prepare the implementation for you.";
            }
        }

        public ImageType Image
        {
            get
            {
                return ImageType.Question;
            }
        }

        public void Translate()
        {
            Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
            if (null != language)
            {
                var component = language.Components["Project Wizard - Environment"];
                Translation.Translator.TranslateControls(this, component.ControlRessources);
            }
            else
            {
                Translation.Translator.TranslateControls(this, "ToolboxControls.ProjectWizard.Controls.EnvironmentControl.txt", Forms.MainForm.Singleton.CurrentLanguageID);
            }
        }

        public void Activate()
        {
             
        }

        public void Deactivate()
        { 
        
        }

        public System.Xml.XmlDocument SettingsDocument
        {
            get { return _settings; }
        }

        public string[] GetSettingsSummary()
        {
            
            string[] result = new string[2];
            result[0] += ProjectWizardControl.Singleton.Localized.Language + Environment.NewLine;
            result[1] += SelectedLanguage + Environment.NewLine;

            result[0] += "IDE" + Environment.NewLine;
            result[1] += "VS " + SelectedIDE + Environment.NewLine;

            result[0] += ProjectWizardControl.Singleton.Localized.Runtime;
            result[1] += SelectedRuntime;

            return result;
        }

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {
            labelNet45Hint.Visible = true;
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

        internal string SelectedLanguage
        { 
            get
            {
                if (radioButtonCSharp.Checked)
                    return "C#";
                else
                    return "VB";
            }
        }

        internal string SelectedIDE
        {
            get
            {
                if (radioButtonVS2010.Checked)
                    return "2010";
                else if (radioButtonVS2012.Checked)
                    return "2012";
                else
                    return "2013";
            }
        }

        internal string SelectedRuntime
        {
            get
            {
                return comboBoxNetRuntime.Text;                 
            }
        }

        private void ChangeSettings()
        {
            _settings.FirstChild.SelectSingleNode("Language").InnerText = SelectedLanguage;
            _settings.FirstChild.SelectSingleNode("IDE").InnerText = SelectedIDE;
            _settings.FirstChild.SelectSingleNode("Runtime").InnerText = SelectedRuntime;
        }

        private void CreateSettingsDocument()
        {
            _settings = new XmlDocument();
            _settings.AppendChild(_settings.CreateElement("EnvironmentControl"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Language"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("IDE"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Runtime"));
        }

        private void RaiseChangeEvent()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        private void radioButtonLanguage_CheckedChanged(object sender, EventArgs e)
        {
            ChangeSettings();
            RaiseChangeEvent();
        }

        private void radioButtonIDE_CheckedChanged(object sender, EventArgs e)
        {
            ChangeSettings();
            RaiseChangeEvent();
        }

        private void comboBoxNetRuntime_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeSettings();
            RaiseChangeEvent();
            if (comboBoxNetRuntime.SelectedIndex == 5)
            {
                // .net 4.5
                labelNet45Hint.Visible = true;
                radioButtonVS2010.Enabled = false;
                radioButtonVS2012.Enabled = false;
                radioButtonVS2013.Enabled = true;
                radioButtonVS2013.Checked = true;
            }
            else
            {
                // else
                labelNet45Hint.Visible = false;
                radioButtonVS2010.Enabled = true;
                radioButtonVS2012.Enabled = true;
                radioButtonVS2013.Enabled = true;
            }
        }
    }
}
