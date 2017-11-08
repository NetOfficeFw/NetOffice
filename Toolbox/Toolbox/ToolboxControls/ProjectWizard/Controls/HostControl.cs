using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    /// <summary>
    /// Supported office applications want selected here
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.HostControl.txt")]
    public partial class HostControl : UserControl, IWizardControl
    {
        #region Fields

        private XmlDocument _settings;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public HostControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
        }

        #endregion

        #region Properties

        /// <summary>
        /// User selected applications
        /// </summary>
        internal List<string> HostApplications
        {
            get 
            {
                List<string> list = new List<string>();
                if (checkBoxExcel.Checked)
                    list.Add("Excel");
                if (checkBoxWord.Checked)
                    list.Add("Word");
                if (checkBoxOutlook.Checked)
                    list.Add("Outlook");
                if (checkBoxPowerPoint.Checked)
                    list.Add("PowerPoint");
                if (checkBoxAccess.Checked)
                    list.Add("Access");
                if (checkBoxProject.Checked)
                    list.Add("Project");
                if (checkBoxVisio.Checked)
                    list.Add("Visio");                
                return list;
            }
        }

        #endregion

        #region IWizardControl Member

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep
        {
            get
            {
                foreach (Control item in this.Controls)
                {
                    CheckBox box = item as CheckBox;
                    if ((null != box) && (box.Checked))
                        return true;
                }
                return false;
            }
        }

        public string Caption
        {
            get
            {
                return "Which Office applications you want support?";
            }
        }

        public string Description
        {
            get
            {
                return "Select one or more Office application(s).";
            }
        }

        public ImageType Image
        {
            get
            {
                return ImageType.Question;
            }
        }

        public XmlDocument SettingsDocument
        {
            get
            {
                return _settings;
            }
        }

        public new void KeyDown(KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.D1:
                    checkBoxExcel.Checked = !checkBoxExcel.Checked;
                    break;
                case Keys.D2:
                    checkBoxWord.Checked = !checkBoxWord.Checked;
                    break;
                case Keys.D3:
                    checkBoxOutlook.Checked = !checkBoxOutlook.Checked;
                    break;
                case Keys.D4:
                    checkBoxPowerPoint.Checked = !checkBoxPowerPoint.Checked;
                    break;
                case Keys.D5:
                    checkBoxAccess.Checked = !checkBoxAccess.Checked;
                    break;
                case Keys.D6:
                    checkBoxProject.Checked = !checkBoxProject.Checked;
                    break;
                case Keys.D7:
                    if (checkBoxVisio.Enabled)
                        checkBoxVisio.Checked = !checkBoxVisio.Checked;
                    break;
                default:
                    break;
            }
        }

        public void Activate()
        {
            foreach (var item in ProjectWizardControl.Singleton.WizardControls)
            {
                ProjectControl ctrl = item as ProjectControl;
                if (null != ctrl)
                {
                    // visio is not supported in NetOffice Tools because it works much different under the hood
                    if (ctrl.IsToolAddinProject)
                    {
                        checkBoxVisio.Checked = false;
                        checkBoxVisio.Enabled = false;
                    }
                    else
                    {
                        checkBoxVisio.Enabled = true;
                    }
                    return;
                }
            }
        }

        public void Deactivate()
        {

        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];
                result[0] = ProjectWizardControl.Singleton.Localized.Applications;

            result[1] = "";

            foreach (XmlNode item in _settings.FirstChild.ChildNodes)
            {
                if (item.Attributes[0].Value.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
                    result[1] += item.Name + " ";
            }

            return result;
        }

        #endregion

        #region Methods

        private void RaiseChangeEvent()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        private void ChangeSettings()
        {
            foreach (Control item in this.Controls)
            {
                CheckBox box = item as CheckBox;
                if (null != box)
                {
                    string name = box.Text.Replace(" ", "");
                    XmlNode node = _settings.FirstChild.SelectSingleNode(name);
                    XmlAttribute attribute = node.Attributes[0];
                    attribute.Value = box.Checked.ToString();
                }
            }
        }

        private void CreateSettingsDocument()
        {
            _settings = new XmlDocument();
            _settings.AppendChild(_settings.CreateElement("Step1Control"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Excel"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Word"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Outlook"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("PowerPoint"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Access"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Project"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Visio"));

            _settings.FirstChild.ChildNodes.Item(0).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(1).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(2).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(3).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(4).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(5).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(6).Attributes.Append(_settings.CreateAttribute("Selected"));
        }

        #endregion

        #region Trigger

        private void checkBox_CheckedChanged(object sender, EventArgs e)
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
       
       #endregion
    }
}
