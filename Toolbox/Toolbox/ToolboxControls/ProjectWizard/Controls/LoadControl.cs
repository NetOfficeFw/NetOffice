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
    /// Loadbehavior in addin projects
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.LoadControl.txt")]
    public partial class LoadControl : UserControl, IWizardControl
    {
        #region Fields

        private XmlDocument _settings;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public LoadControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
            comboBoxLoadBehavior.SelectedIndex = 0;
        }

        #endregion

        /// <summary>
        /// LoadBehavior Registry Value
        /// </summary>
        internal string LoadBehaviour
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

        /// <summary>
        /// Addin Root RegistryKey (Current user or local machine)
        /// </summary>
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
                return "Choose Load Behavior options.";
            }
        }

        public string Description
        {
            get
            {
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

        public new void KeyDown(KeyEventArgs e)
        {

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
                result[0] += ProjectWizardControl.Singleton.Localized.Registry;
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

        #region Methods
         
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

        #region Trigger

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                ChangeSettings();
                RaiseChangeEvent();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void comboBoxLoadBehavior_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //if (noChangeEventFlag)
                //    return;

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
