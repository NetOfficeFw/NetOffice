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
    /// <summary>
    /// User interface options in addin projects
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.GuiControl.txt")]
    public partial class GuiControl : UserControl, IWizardControl
    {
        #region Fields

        private XmlDocument _settings;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public GuiControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
        }

        #endregion

        #region Properties

        /// <summary>
        /// User want support classic ui
        /// </summary>
        internal bool ClassicUIEnabled
        {
            get 
            {
                return checkBoxClassicUISupport.Checked;
            }
        }

        /// <summary>
        /// User want support ugly ui
        /// </summary>
        internal bool RibbonUIEnabled
        {
            get
            {
                return checkBoxRibbonUISupport.Checked;
            }
        }

        /// <summary>
        /// User want support a task pane
        /// </summary>
        internal bool TaskPaneEnabled
        {
            get
            {
                return checkBoxTaskPaneSupport.Checked;
            }
        }

        /// <summary>
        /// User want a ribbon toogle button for taskpane visibilty
        /// </summary>
        internal bool ToogleEnabled
        {
            get
            {
                return checkBoxToogleButton.Checked && checkBoxToogleButton.Enabled; 
            }
        }

        private string LocalizedUglyUI
        {
            get
            {
                return "Ribbon UI:";
            }
        }

        private string LocalizedClassicUI
        {
            get
            {
                return "Classic UI:";
            }
        }

        #endregion

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
                return "User Interface";
            }
        }

        public string Description
        {
            get
            {
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
                result[1] += ProjectWizardControl.Singleton.Localized.Yes;
            }
            else
            {
                result[0] += LocalizedClassicUI;
                result[1] += ProjectWizardControl.Singleton.Localized.No;
            }

            if (_settings.FirstChild.SelectSingleNode("UseRibbonUI").InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
            {
                result[0] += Environment.NewLine + LocalizedUglyUI;
                result[1] += Environment.NewLine + ProjectWizardControl.Singleton.Localized.Yes;
            }
            else
            {
                result[0] += Environment.NewLine + LocalizedUglyUI;
                result[1] += Environment.NewLine + ProjectWizardControl.Singleton.Localized.No;
            }
           
            string toogle = true == checkBoxToogleButton.Checked ? " + Toogle" : "";
            result[0] += Environment.NewLine + "TaskPane";
            if (checkBoxTaskPaneSupport.Checked)
                result[1] += Environment.NewLine + ProjectWizardControl.Singleton.Localized.Yes + toogle;
            else
                result[1] += Environment.NewLine + ProjectWizardControl.Singleton.Localized.No;

            return result;
        }

        #endregion

        #region Methods
       
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

        #region Trigger

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                checkBoxToogleButton.Enabled = checkBoxTaskPaneSupport.Checked && checkBoxRibbonUISupport.Checked;
                if (!checkBoxToogleButton.Enabled && checkBoxToogleButton.Checked)
                    checkBoxToogleButton.Checked = false;
                ChangeSettings();
                RaiseChangeEvent();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            } 
        }

        private void checkBoxTaskPaneSupport_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                checkBoxToogleButton.Enabled = checkBoxTaskPaneSupport.Checked && checkBoxRibbonUISupport.Checked;
                if (!checkBoxToogleButton.Enabled && checkBoxToogleButton.Checked)
                    checkBoxToogleButton.Checked = false;
                ChangeSettings();
                RaiseChangeEvent();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            } 
        }

        #endregion
    }
}