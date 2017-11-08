using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Xml;
using System.Text;
using System.Windows.Forms;
using System.Security.Principal;
using Microsoft.Win32;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    /// <summary>
    /// Project type want selected here as first
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.ProjectControl.txt")]
    public partial class ProjectControl : UserControl, IWizardControl
    {
        #region Fields

        private XmlDocument _settings;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ProjectControl()
        {
            InitializeComponent(); 
            CreateSettingsDocument();
            if (!Program.IsAdmin)
                labelNoAdminHint.Visible = true;
            else
                labelNoAdminHint.Visible = false;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Returns project is a NetOffice Tools addin project
        /// </summary>
        internal bool IsToolAddinProject
        {
            get
            {
                return checkBoxUseTools.Checked;
            }
        }

        /// <summary>
        /// Returns selected output folder
        /// </summary>
        internal string CalculatedFolder
        {
            get
            {
                if (radioButtonCustomFolder.Checked)
                    return textBoxCustomFolder.Text;
                else
                {
                    if (radioButtonDesktop.Checked)
                        return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    if (radioButtonUserFolder.Checked)
                        return Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                    if (radioButtonApplicationData.Checked)
                        return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    return ProjectOptions.GetVisualStudioProjectFolder();
                }
            }
        }

        /// <summary>
        /// User want use the NetOffice tools
        /// </summary>
        internal bool UseTools
        {
            get
            {
                if (checkBoxUseTools.Checked && radioButtonAutomationAddin.Checked)
                    return true;
                else
                    return false;
            }
        }

        /// <summary>
        /// Current selected folder kind
        /// </summary>
        internal string SelectedFolder
        {
            get
            {
                switch (SelectedProjectFolderType())
                {
                    case "Custom":
                        return textBoxCustomFolder.Text;
                    default:
                        return "";
                }
            }

        }
      
        #endregion

        #region IWizardControl

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep
        {
            get
            {
                if (radioButtonCustomFolder.Checked && string.IsNullOrWhiteSpace(textBoxCustomFolder.Text))
                    return false;
                else
                    return true; 
            }
        }

        public string Caption
        {
            get
            {
                return "Project Type";
            }
        }

        public string Description
        {
            get
            {
                return "Select your project type.";
            }
        }

        public ImageType Image
        {
            get
            {
                return ImageType.Question;
            }
        }

        public void Activate()
        {
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

        public new void KeyDown(KeyEventArgs e)
        {

        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];

            result[0] += ProjectWizardControl.Singleton.Localized.ProjectType + Environment.NewLine;
            result[1] += SelectedProjectType() + Environment.NewLine;

            result[0] += ProjectWizardControl.Singleton.Localized.ProjectFolder;
            result[1] += SelectedProjectFolderType();

            return result;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Current selected project type
        /// </summary>
        /// <returns></returns>
        internal ProjectType SelectedProjectType()
        {

            if (radioButtonAutomationAddin.Checked)
            {
                if (checkBoxUseTools.Checked)
                    return ProjectType.NetOfficeAddin;
                else
                    return ProjectType.NetOfficeAddin;
            }
            if (radioButtonWindowsForms.Checked)
                return ProjectType.WindowsForms;
            if (radioButtonConsole.Checked)
                return ProjectType.Console;
            if (radioButtonClassLibrary.Checked)
                return ProjectType.ClassLibrary;
            throw new IndexOutOfRangeException("SelectedProjectType");
        }

        /// <summary>
        /// Current selected project folder type
        /// </summary>
        /// <returns></returns>
        internal string SelectedProjectFolderType()
        {
            if (radioButtonApplicationData.Checked)
                return radioButtonApplicationData.Text;
            if (radioButtonDesktop.Checked)
                return radioButtonDesktop.Text;
            if (radioButtonUserFolder.Checked)
                return radioButtonUserFolder.Text;
            if (radioButtonVSProjectFolder.Checked)
                return radioButtonVSProjectFolder.Text;
            if (radioButtonCustomFolder.Checked)
                return radioButtonVSProjectFolder.Text;
            throw new IndexOutOfRangeException("SelectedProjectFolderType");
        }

        private void ChangeSettings()
        {
            _settings.FirstChild.SelectSingleNode("ProjectType").InnerText = SelectedProjectType().ToString();
            _settings.FirstChild.SelectSingleNode("ProjectFolderType").InnerText = SelectedProjectFolderType();
            _settings.FirstChild.SelectSingleNode("ProjectFolder").InnerText = SelectedFolder;
        }

        private void CreateSettingsDocument()
        {
            _settings = new XmlDocument();
            _settings.AppendChild(_settings.CreateElement("ProjectControl"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("ProjectType"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("ProjectFolderType"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("ProjectFolder"));
        }

        private void RaiseChangeEvent()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        #endregion

        #region Trigger

        private void radioButtonProjectType_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                checkBoxUseTools.Enabled = radioButtonAutomationAddin.Checked;
                checkBoxUseTools.Checked = radioButtonAutomationAddin.Checked;
                ChangeSettings();
                RaiseChangeEvent();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void radioButtonProjectFolder_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                ChangeSettings();
                RaiseChangeEvent();

                RadioButton button = sender as RadioButton;
                buttonChooseFolder.Enabled = radioButtonCustomFolder.Checked;
                if (radioButtonCustomFolder.Checked && string.IsNullOrWhiteSpace(textBoxCustomFolder.Text) & button == radioButtonCustomFolder)
                    buttonChooseFolder_Click(buttonChooseFolder, new EventArgs());
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                if (DialogResult.OK == dialog.ShowDialog(this))
                {
                    textBoxCustomFolder.Text = dialog.SelectedPath;
                    ChangeSettings();
                    RaiseChangeEvent();
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }       

        #endregion
    }
}