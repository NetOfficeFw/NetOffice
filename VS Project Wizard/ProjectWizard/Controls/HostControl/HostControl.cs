using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.ProjectWizard
{
    internal partial class HostControl : UserControl, IWizardControl 
    {
        XmlDocument _settings;

        public HostControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
        }
         
        #region IWizardControl Member

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep 
        { 
            get
            {
                foreach (Control item in this.Controls)
                {
                    CheckBox box = item as CheckBox;
                    if((null != box) && (box.Checked))
                        return true;
                }
                return false;
            }
        }

        public string Caption
        {
            get
            {
                if (NetOfficeProject.TargetLanguage == TargetLanguage.German)
                    return "Welche Office Anwendungen möchten Sie unterstützen?";
                else
                    return "Which Office applications you want support?";
            }
        }

        public string Description 
        {
            get
            {
                if (NetOfficeProject.TargetLanguage == TargetLanguage.German)
                    return "Wählen Sie eine oder mehrere Office Anwendungen.";
                else
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

        public void Translate()
        {

        }

        public void Activate()
        {

        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];
            if (NetOfficeProject.TargetLanguage == TargetLanguage.German)
                result[0] = "Office Anwendungen:";
             else
                 result[0] = "Office Applications:";

            result[1] = "";

            foreach (XmlNode item in _settings.FirstChild.ChildNodes)
            {
                if(item.Attributes[0].Value.Equals("TRUE",StringComparison.InvariantCultureIgnoreCase))
                    result[1] += item.Name + " ";
            }

            return result;
        }

        #endregion

        #region Trigger

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            ChangeSettings();
            RaiseChangeEvent();
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

            _settings.FirstChild.ChildNodes.Item(0).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(1).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(2).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(3).Attributes.Append(_settings.CreateAttribute("Selected"));
            _settings.FirstChild.ChildNodes.Item(4).Attributes.Append(_settings.CreateAttribute("Selected"));
        }

        #endregion
    }
}
