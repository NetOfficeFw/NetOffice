using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    public partial class GuiControl : UserControl, IWizardControl
    {
        XmlDocument _settings;

        public GuiControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
            if (ProjectWizardControl.CurrentLanguageID == 1031)
            {
                checkBoxClassicUISupport.Text = "Ich möchte die klassische Benutzeroberfläche in älteren Office Versionen erweitern";
                checkBoxRibbonUISupport.Text = "Ich möchte die Ribbon Oberfläche in neueren Office Versionen erweitern";
            }
            else
            {
                checkBoxClassicUISupport.Text = "Ich want customize the classic User Interface.";
                checkBoxRibbonUISupport.Text = "Ich want customize the Ribbon User Interface.";
            }
        }

        #region IWizardControl Member

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep
        {
            get { return true; }
        }

        public string Caption
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Benutzerschnittstelle";
                else
                    return "User Interface";
            }
        }

        public string Description
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
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
        Translator.TranslateControls(this, "ProjectWizard.Controls.GuiControl.txt", ProjectWizardControl.CurrentLanguageID);
        }

        public void Activate()
        {
            ChangeSettings();
            RaiseChangeEvent();
        }

        public System.Xml.XmlDocument SettingsDocument
        {
            get { return _settings; }
        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];
            result[0] = "";
            result[1] = "";

            if (_settings.FirstChild.SelectSingleNode("UseClassicUI").InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
            {
                result[0] += LocalizedClassicUI;
                result[1] += LocalizedYes;
            }
            else
            {
                result[0] += LocalizedClassicUI;
                result[1] += LocalizedNo;
            }

            if (_settings.FirstChild.SelectSingleNode("UseRibbonUI").InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
            {
                result[0] += Environment.NewLine + LocalizedUglyUI;
                result[1] += Environment.NewLine + LocalizedYes;
            }
            else
            {
                result[0] += Environment.NewLine + LocalizedUglyUI;
                result[1] += Environment.NewLine + LocalizedNo;
            }

            return result;
        }

        #endregion

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            ChangeSettings();
            RaiseChangeEvent();
        }

        #region Methods

        private string LocalizedUglyUI
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Ribbon UI:";
                else
                    return "Ribbon UI:";
            }
        }

        private string LocalizedClassicUI
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Klassische UI:";
                else
                    return "Classic UI:";
            }
        }

        private string LocalizedYes
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Ja";
                else
                    return "Yes";
            }
        }

        private string LocalizedNo
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Nein";
                else
                    return "No";
            }
        }

        private void ChangeSettings()
        {
            _settings.FirstChild.SelectSingleNode("UseClassicUI").InnerText = checkBoxClassicUISupport.Checked.ToString();
            _settings.FirstChild.SelectSingleNode("UseRibbonUI").InnerText = checkBoxRibbonUISupport.Checked.ToString();
        }

        private void CreateSettingsDocument()
        {
            _settings = new XmlDocument();
            _settings.AppendChild(_settings.CreateElement("Step4Control"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("UseClassicUI"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("UseRibbonUI"));
        }

        private void RaiseChangeEvent()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        #endregion   
    }
}
