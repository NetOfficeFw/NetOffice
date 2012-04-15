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
    public partial class LoadControl : UserControl, IWizardControl
    {
        XmlDocument _settings;
        bool noChangeEventFlag;

        public LoadControl()
        {
            InitializeComponent();
            Translate();
            CreateSettingsDocument();
            comboBoxLoadBehavior.SelectedIndex = 0;
        }
        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            ChangeSettings();
            RaiseChangeEvent();
        }

        private void comboBoxLoadBehavior_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (noChangeEventFlag)
                return;

            ChangeSettings();
            RaiseChangeEvent();
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
                    return "Vom wem soll Ihr Automations Addin geladen werden?";
                else
                    return "Choose Load Behavior options.";
            }
        }

        public string Description
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Entscheiden Sie wann und von wem Ihr Addin geladen wird.";
                else
                    return "Decide when it has to be loaded";
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
            noChangeEventFlag = true;
            int selIndex = comboBoxLoadBehavior.SelectedIndex;

            Translator.TranslateControls(this, "ProjectWizard.Controls.LoadControl.txt", ProjectWizardControl.CurrentLanguageID);

            comboBoxLoadBehavior.SelectedIndex = selIndex;
            noChangeEventFlag = false;
        }

        public void Activate()
        {
            ChangeSettings();
            RaiseChangeEvent();
        }

        public XmlDocument SettingsDocument
        {
            get
            {
                return _settings;
            }
        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];
            result[0] = "";
            result[1] = "";

            if (_settings.FirstChild.ChildNodes[0].InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
            {
                result[0] += LocalizedRegistry;
                result[1] += LocalizedRegistryModeMachine;
            }
            else
            {
                result[0] += LocalizedRegistry;
                result[1] += LocalizedRegistryModeCurrentUser;
            }

            result[0] += Environment.NewLine + LocalizedLoadBehavior;
            result[1] += Environment.NewLine + TranslateLoadBehavior(_settings.FirstChild.ChildNodes[1].InnerText);

            return result;
        }

        #endregion

        #region Methods

        private string LocalizedLoadBehavior
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Ladeverhalten:";
                else
                    return "Load Behavior:";
            }
        }

        private string LocalizedRegistryModeMachine
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Für alle Benutzer";
                else
                    return "All Users";
            }
        }

        private string LocalizedRegistryModeCurrentUser
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Nur für den aktuellen Benutzer";
                else
                    return "Current User";
            }
        }

        private string LocalizedRegistry
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Registrierung:";
                else
                    return "Registry:";
            }
        }

        private string TranslateLoadBehavior(string value)
        {
            switch (value)
            {
                case "3":
                    if (ProjectWizardControl.CurrentLanguageID == 1031)
                        return "Beim Start der Office Anwendung";
                    else
                        return "Load at startup";
                case "2":
                    if (ProjectWizardControl.CurrentLanguageID == 1031)
                        return "Bei Bedarf";
                    else
                        return "On Demand";
                case "1":
                    if (ProjectWizardControl.CurrentLanguageID == 1031)
                        return "Nicht automatisch laden";
                    else
                        return "Not automatically";
                case "16":
                    if (ProjectWizardControl.CurrentLanguageID == 1031)
                        return "Beim ersten Start automatisch, danach bei Bedarf";
                    else
                        return "Load first time, then load on demand";
                default:
                    throw new ArgumentOutOfRangeException("TranslateLoadBehavior:value");
            }
        }

        private void ChangeSettings()
        {
            _settings.FirstChild.SelectSingleNode("RegisterHKeyLocalMachine").InnerText = radioButtonLocalMachine.Checked.ToString();

            string value = "";
            switch (comboBoxLoadBehavior.SelectedIndex)
            {
                case 0:
                    value = "3";
                    break;
                case 1:
                    value = "2";
                    break;
                case 2:
                    value = "1";
                    break;
                case 3:
                    value = "16";
                    break;
                default:
                    throw new ArgumentOutOfRangeException("comboBoxLoadBehavior.SelectedIndex");
            }
            _settings.FirstChild.SelectSingleNode("LoadBehavior").InnerText = value;
        }

        private void CreateSettingsDocument()
        {
            _settings = new XmlDocument();
            _settings.AppendChild(_settings.CreateElement("Step3Control"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("RegisterHKeyLocalMachine"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("LoadBehavior"));
            _settings.FirstChild.ChildNodes[1].InnerText = "3";
        }

        private void RaiseChangeEvent()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        #endregion
    }
}
