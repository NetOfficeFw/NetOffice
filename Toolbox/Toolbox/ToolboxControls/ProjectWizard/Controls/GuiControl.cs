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
    public partial class GuiControl : UserControl, IWizardControl
    {
        XmlDocument _settings;

        public GuiControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
            if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
            {
                checkBoxClassicUISupport.Text = "Ich möchte die klassische Benutzeroberfläche in älteren Office Versionen erweitern";
                checkBoxRibbonUISupport.Text = "Ich möchte die Ribbon Oberfläche in neueren Office Versionen erweitern";
                checkBoxTaskPaneSupport.Text = "Ich möchte eine Task Pane zur Verfügung stellen";
            }
            else
            {
                checkBoxClassicUISupport.Text = "I want customize the classic User Interface";
                checkBoxRibbonUISupport.Text = "I want customize the Ribbon User Interface";
                checkBoxTaskPaneSupport.Text = "I want a custom Task Pane";
            }
        }

        public bool ClassicUIEnabled
        {
            get 
            {
                return checkBoxClassicUISupport.Checked;
            }
        }

        public bool RibbonUIEnabled
        {
            get
            {
                return checkBoxRibbonUISupport.Checked;
            }
        }

        public bool ToogleEnabled
        {
            get
            {
                return checkBoxToogleButton.Checked && checkBoxToogleButton.Enabled; 
            }
        }

        public bool TaskPaneEnabled
        {
            get
            {
                return checkBoxTaskPaneSupport.Checked;
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
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
                    return "Benutzerschnittstelle";
                else
                    return "User Interface";
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
            Translation.Translator.TranslateControls(this, "ToolboxControls.ProjectWizard.Controls.GuiControl.txt", ProjectWizardControl.Singleton.Host.CurrentLanguageID);
        }

        public new void KeyDown(KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.D1:
                    checkBoxClassicUISupport.Checked = !checkBoxClassicUISupport.Checked;
                    break;
                case Keys.D2:
                    checkBoxRibbonUISupport.Checked = !checkBoxRibbonUISupport.Checked;
                    break;
                case Keys.D3:
                    checkBoxTaskPaneSupport.Checked = !checkBoxTaskPaneSupport.Checked;
                    break;
                case Keys.D4:
                    if(checkBoxToogleButton.Enabled)
                        checkBoxToogleButton.Checked = !checkBoxToogleButton.Checked;
                    break;
                default:
                    break;
            }
        }

        public void Activate()
        {
            if (ProjectWizardControl.Singleton.IsSingleVisioProject)
            {
                checkBoxClassicUISupport.Checked = false;
                checkBoxRibbonUISupport.Checked = false;
                checkBoxTaskPaneSupport.Checked = false;
                checkBoxClassicUISupport.Enabled = false;
                checkBoxRibbonUISupport.Enabled = false;
                checkBoxTaskPaneSupport.Enabled = false;
            }
            else
            {
                checkBoxClassicUISupport.Enabled = true;
                checkBoxRibbonUISupport.Enabled = true;
                checkBoxTaskPaneSupport.Enabled = true;
            }

            checkBoxToogleButton.Visible = !ProjectWizardControl.Singleton.IsSimpleAddinProject;
            checkBoxToogleButton.Enabled = false == ProjectWizardControl.Singleton.IsSimpleAddinProject && checkBoxTaskPaneSupport.Checked && checkBoxRibbonUISupport.Checked;

            ChangeSettings();
            RaiseChangeEvent();
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
            
            string toogle = true == checkBoxToogleButton.Checked ? " + Toogle" : "";
            result[0] += Environment.NewLine + "TaskPane";
            if (checkBoxRibbonUISupport.Checked)
                result[1] += Environment.NewLine + LocalizedYes + toogle;
            else
                result[1] += Environment.NewLine + LocalizedNo;

            return result;
        }

        #endregion

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxToogleButton.Enabled = checkBoxTaskPaneSupport.Checked && checkBoxRibbonUISupport.Checked;
            if (!checkBoxToogleButton.Enabled && checkBoxToogleButton.Checked)
                checkBoxToogleButton.Checked = false;
            ChangeSettings();
            RaiseChangeEvent();
        }

        private void checkBoxTaskPaneSupport_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxToogleButton.Enabled = checkBoxTaskPaneSupport.Checked && checkBoxRibbonUISupport.Checked;
            if (!checkBoxToogleButton.Enabled && checkBoxToogleButton.Checked)
                checkBoxToogleButton.Checked = false;
            ChangeSettings();
            RaiseChangeEvent();
        }

        #region Methods

        private string LocalizedUglyUI
        {
            get
            {
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
                    return "Ribbon UI:";
                else
                    return "Ribbon UI:";
            }
        }

        private string LocalizedClassicUI
        {
            get
            {
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
                    return "Klassische UI:";
                else
                    return "Classic UI:";
            }
        }

        private string LocalizedYes
        {
            get
            {
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
                    return "Ja";
                else
                    return "Yes";
            }
        }

        private string LocalizedNo
        {
            get
            {
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
                    return "Nein";
                else
                    return "No";
            }
        }

         
        private void ChangeSettings()
        {
            _settings.FirstChild.SelectSingleNode("UseClassicUI").InnerText = checkBoxClassicUISupport.Checked.ToString();
            _settings.FirstChild.SelectSingleNode("UseRibbonUI").InnerText = checkBoxRibbonUISupport.Checked.ToString();
            _settings.FirstChild.SelectSingleNode("UseTaskPane").InnerText = checkBoxTaskPaneSupport.Checked.ToString();
            _settings.FirstChild.SelectSingleNode("UseToogle").InnerText = checkBoxToogleButton.Checked.ToString();
        }

        private void CreateSettingsDocument()
        {
            _settings = new XmlDocument();
            _settings.AppendChild(_settings.CreateElement("Step4Control"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("UseClassicUI"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("UseRibbonUI"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("UseTaskPane"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("UseToogle"));
        }

        private void RaiseChangeEvent()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        #endregion   
    }
}
