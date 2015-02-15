using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.LoadControl.txt")]
    public partial class LoadControl : UserControl, IWizardControl, ILocalizationDesign
    {
        XmlDocument _settings;
        bool noChangeEventFlag;

        public LoadControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
            comboBoxLoadBehavior.SelectedIndex = 0;
        }
       
        public new void KeyDown(KeyEventArgs e)
        {

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

        public string LoadBehaviour
        {
            get
            {
                switch (comboBoxLoadBehavior.SelectedIndex)
                {
                    case 0:
                        return "3";
                    case 1:
                        return "2";
                    case 2:
                        return "1";
                    case 3:
                        return "16";
                    default:
                        return "3";
                }
            }
        }

        public string Hivekey
        {
            get 
            {
                if (radioButtonCurrentUser.Checked)
                    return "CurrentUser";
                else
                    return "LocalMachine";
            }
        }

        #region IWizardControl

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep
        {
            get { return true; }
        }

        public string Caption
        {
            get
            {
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
                    return "Vom wem soll Ihr Automations Addin geladen werden?";
                else
                    return "Choose Load Behavior options.";
            }
        }

        public string Description
        {
            get
            {
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
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
            Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
            if (null != language)
            {
                var component = language.Components["Project Wizard - Load"];
                Translation.Translator.TranslateControls(this, component.ControlRessources);
            }
            else
            {
                Translation.Translator.TranslateControls(this, "ToolboxControls.ProjectWizard.Controls.LoadControl.txt", Forms.MainForm.Singleton.CurrentLanguageID);
            }
            TranslateComboBox();
            comboBoxLoadBehavior.SelectedIndex = selIndex;
            noChangeEventFlag = false;
        }

        private void TranslateComboBox()
        {
            comboBoxLoadBehavior.Items.Clear();
            comboBoxLoadBehavior.Items.Add(ProjectWizardControl.Singleton.Localized.AddinStartup);
            comboBoxLoadBehavior.Items.Add(ProjectWizardControl.Singleton.Localized.AddinOnDemand);
            comboBoxLoadBehavior.Items.Add(ProjectWizardControl.Singleton.Localized.AddinNotAutomaticaly);
            comboBoxLoadBehavior.Items.Add(ProjectWizardControl.Singleton.Localized.AddinFirstTime);
        }

        public void Activate()
        {
            ChangeSettings();
            RaiseChangeEvent();
        }

        public void Deactivate()
        {

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
                result[0] += ProjectWizardControl.Singleton.Localized.Registry; ;
                result[1] += ProjectWizardControl.Singleton.Localized.RegistryLocalMachine;
            }
            else
            {
                result[0] += ProjectWizardControl.Singleton.Localized.Registry;
                result[1] += ProjectWizardControl.Singleton.Localized.RegistryCurrentUser;
            }

            result[0] += Environment.NewLine + ProjectWizardControl.Singleton.Localized.LoadBehavior;
            result[1] += Environment.NewLine + TranslateLoadBehavior();

            return result;
        }

        #endregion

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {
            //comboBoxLoadBehavior.Visible = false;
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
         
        //private string LocalizedLoadBehavior
        //{
        //    get
        //    {
        //        if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
        //            return "Ladeverhalten:";
        //        else
        //            return "Load Behavior:";
        //    }
        //}

        //private string LocalizedRegistryModeMachine
        //{
        //    get
        //    {
        //        if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
        //            return "Für alle Benutzer";
        //        else
        //            return "All Users";
        //    }
        //}

        //private string LocalizedRegistryModeCurrentUser
        //{
        //    get
        //    {
        //        if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
        //            return "Nur für den aktuellen Benutzer";
        //        else
        //            return "Current User";
        //    }
        //}

        //private string LocalizedRegistry
        //{
        //    get
        //    {
        //        if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
        //            return "Registrierung:";
        //        else
        //            return "Registry:";
        //    }
        //}

        private string TranslateLoadBehavior()
        {
            int index = comboBoxLoadBehavior.Text.IndexOf("=");
            if (index > -1)
            {
                string text = comboBoxLoadBehavior.Text.Substring(index + 1).Trim();
                return text;
            }
            else
                return comboBoxLoadBehavior.Text;
        }

        //private string TranslateLoadBehavior(string value)
        //{
        //    switch (value)
        //    {
        //        case "3":
        //            if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
        //                return "Beim Start der Office Anwendung";
        //            else
        //                return "Load at startup";
        //        case "2":
        //            if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
        //                return "Bei Bedarf";
        //            else
        //                return "On Demand";
        //        case "1":
        //            if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
        //                return "Nicht automatisch laden";
        //            else
        //                return "Not automatically";
        //        case "16":
        //            if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
        //                return "Beim ersten Start automatisch, danach bei Bedarf";
        //            else
        //                return "Load first time, then load on demand";
        //        default:
        //            throw new ArgumentOutOfRangeException("TranslateLoadBehavior:value");
        //    }
        //}

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
