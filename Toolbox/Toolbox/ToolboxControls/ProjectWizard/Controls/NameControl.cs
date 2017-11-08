using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    /// <summary>
    /// Name and its description for the target assembly
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.NameControl.txt")]
    public partial class NameControl : UserControl, IWizardControl
    {
        #region Fields

        private XmlDocument _settings;
        private bool _firstActivateFlag;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public NameControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
            ChangeSettings();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Name for the assembly
        /// </summary>
        internal string AssemblyName
        {
            get { return textBoxClassName.Text.Trim(); }
        }

        /// <summary>
        /// Description for the assembly
        /// </summary>
        internal string AssemblyDescription
        {
            get { return textBoxDescription.Text; }
        }

        #endregion

        #region IWizardControl

        public new void KeyDown(KeyEventArgs e)
        {

        }

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep
        {
            get
            {
                try
                {
                    if(null != ClassNameInvalid())
                        return false;
                    return (("" != textBoxClassName.Text.Trim()) && (!ProjectWizardControl.Singleton.FolderExists(textBoxClassName.Text.Trim())));
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
                return "Informations about your assembly.";
            }
        }

        public string Description
        {
            get
            {
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

        public void Activate()
        {
            textBoxClassName.Focus();
            if (_firstActivateFlag == false)
            {
                SetEnglishDefaultName();
                _firstActivateFlag = true;
            }
        }

        public void Deactivate()
        {

        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];
            result[0] = "";
            result[1] = "";

            string name = _settings.FirstChild.ChildNodes[0].InnerText;
            string description = _settings.FirstChild.ChildNodes[1].InnerText;

            result[0] += labelClassName.Text + Environment.NewLine + labelDescription.Text;
            result[1] += name + Environment.NewLine + description;

            return result;
        }

        #endregion

        #region Methods

        private void SetEnglishDefaultName()
        {
            if (!ProjectWizardControl.Singleton.FolderExists("MyAssembly"))
            {
                textBoxClassName.Text = "MyAssembly";
                textBoxDescription.Text = "Assembly Description";
            }
            else
            {
                for (int i = 2; i < 999; i++)
                {
                    string name = "MyAssembly" + i.ToString();
                    if (!ProjectWizardControl.Singleton.FolderExists(name))
                    {
                        textBoxClassName.Text = name;
                        textBoxDescription.Text = "Assembly Description";
                        return;
                    }
                }
            }
        }

        private void SetGermanDefaultName()
        {
            if (!ProjectWizardControl.Singleton.FolderExists("MeinAssembly"))
            {
                textBoxClassName.Text = "MeinAssembly";
                textBoxDescription.Text = "Assembly Beschreibung";
            }
            else
            {
                for (int i = 2; i < 999; i++)
                {
                    string name = "MeinAssembly" + i.ToString();
                    if (!ProjectWizardControl.Singleton.FolderExists(name))
                    {
                        textBoxClassName.Text = name;
                        textBoxDescription.Text = "Assembly Beschreibung";
                        return;
                    }
                }
            }
        }

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
       
        private string ClassNameInvalid()
        {
            string text = textBoxClassName.Text;
            foreach (var item in System.IO.Path.GetInvalidFileNameChars())
            {
                string str = item.ToString();
                if (str.Trim().Length == 0)
                    continue;
                if (textBoxClassName.Text.IndexOf(item) > -1)
                    return item.ToString();
            }
            foreach (var item in System.IO.Path.GetInvalidPathChars())
            {
                string str = item.ToString();
                if (str.Trim().Length == 0)
                    continue;
                if (textBoxClassName.Text.IndexOf(item) > -1)
                    return item.ToString();
            }
            return null;
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

        #endregion

        #region Trigger

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ChangeSettings();
                RaiseChangeEvent();
                if (ProjectWizardControl.Singleton.FolderExists(textBoxClassName.Text.Trim()))
                    labelHint.Visible = true;
                else
                    labelHint.Visible = false;

                string res = ClassNameInvalid();
                if (null != res)
                {
                    errorProvider1.SetError(textBoxClassName, "Invalid Char: " + res);
                    return;
                }

                if (textBoxClassName.Text.IndexOf(" ") > -1)
                {
                    errorProvider1.SetError(textBoxClassName, "Invalid empty space");
                    return;
                }

                errorProvider1.Clear();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
