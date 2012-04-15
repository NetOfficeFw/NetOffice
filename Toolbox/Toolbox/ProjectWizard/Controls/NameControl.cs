using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    public partial class NameControl : UserControl, IWizardControl
    {
        XmlDocument _settings;

        public NameControl()
        {
            InitializeComponent();
            Translate();
            textBoxClassName.Text = "MyAssembly";
            textBoxDescription.Text = "No Description available";
            CreateSettingsDocument();
            ChangeSettings();
        }

        #region IWizardControl Member

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep
        {
            get
            {
                try
                {
                    return (("" != textBoxClassName.Text.Trim()) || ("" != textBoxClassName.Text.Trim()));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("RaiseChangeEvent " + ex.Message);
                    throw (ex);
                }
            }
        }

        public string Caption
        {
            get
            {

                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Tragen Sie Informationen zu Ihrem Assembly ein.";
                else
                    return "Informations about your assembly.";
            }
        }

        public string Description
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Diese Informationen sind für Anwender sichtbar.";
                else
                    return "These informations are visible for your customers.";
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
            Translator.TranslateControls(this, "ProjectWizard.Controls.NameControl.txt", ProjectWizardControl.CurrentLanguageID);
        }

        public void Activate()
        {
            textBoxClassName.Focus();
        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];
            result[0] = "";
            result[1] = "";

            string name = _settings.FirstChild.ChildNodes[0].InnerText;
            string description = _settings.FirstChild.ChildNodes[1].InnerText;

            if (ProjectWizardControl.CurrentLanguageID == 1031)
                result[0] += "Assembly Name:" + Environment.NewLine + "Assembly Beschreibung:";
            else
                result[0] += "Assembly Name:" + Environment.NewLine + "Assembly Description:";

            result[1] += name + Environment.NewLine + description;

            return result;
        }

        #endregion

        #region Methods

        private void RaiseChangeEvent()
        {
            try
            {
                if (null != ReadyStateChanged)
                    ReadyStateChanged(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RaiseChangeEvent " + ex.Message);
            }
        }

        #endregion

        private void textBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void ChangeSettings()
        {
            foreach (Control item in this.Controls)
            {
                TextBox box = item as TextBox;
                if (null != box)
                {
                    string name = box.Name.Substring("textBox".Length);
                    XmlNode node = _settings.FirstChild.SelectSingleNode(name);
                    if (box.Name == "textBoxClassName")
                        node.InnerText = box.Text.Trim().Replace(" ", "");
                    else
                        node.InnerText = box.Text.Trim();
                }
            }
        }

        private void CreateSettingsDocument()
        {
            _settings = new XmlDocument();
            _settings.AppendChild(_settings.CreateElement("Step2Control"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("ClassName"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Description"));
        }
    }
}
