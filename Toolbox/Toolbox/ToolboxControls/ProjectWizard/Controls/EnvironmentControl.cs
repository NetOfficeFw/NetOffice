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
    public partial class EnvironmentControl : UserControl, IWizardControl
    {
        XmlDocument _settings;

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
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
                    return "Umgebung";
                else
                    return "Environment";
            }
        }

        public string Description
        {
            get
            {
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
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
            Translation.Translator.TranslateControls(this, "ToolboxControls.ProjectWizard.Controls.EnvironmentControl.txt", ProjectWizardControl.Singleton.Host.CurrentLanguageID);
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
            result[0] += ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1033 ? "Language" + Environment.NewLine : "Sprache" + Environment.NewLine;
            result[1] += ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1033 ? SelectedLanguage + Environment.NewLine : SelectedLanguage + Environment.NewLine;

            result[0] += ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1033 ? "IDE" + Environment.NewLine : "IDE" + Environment.NewLine;
            result[1] += ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1033 ? "VS " + SelectedIDE + Environment.NewLine : "VS " + SelectedIDE + Environment.NewLine;

            result[0] += ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1033 ? ".NET Runtime" : ".NET Lautzeitumgebung";
            result[1] += ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1033 ? SelectedRuntime : SelectedRuntime;

            return result;
        }

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
