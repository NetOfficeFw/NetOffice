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

namespace NetOffice.DeveloperToolbox
{
    public partial class ProjectControl : UserControl, IWizardControl
    {
        XmlDocument _settings;

        public ProjectControl()
        {
            InitializeComponent(); 
            CreateSettingsDocument();
            if (!IsAdministrator())
            {
                labelNoAdminHint.Visible = true;
            }
            else
            {
                labelNoAdminHint.Visible = false;
            }
        }

        #region IWizardControl Member

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
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Project Typ";
                else
                    return "Project Type";
            }
        }

        public string Description
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Was für ein Projekt soll erstellt werden?";
                else
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

        public void Translate()
        {
            Translator.TranslateControls(this, "ProjectWizard.Controls.ProjectControl.txt", ProjectWizardControl.CurrentLanguageID);
        }

        public void Activate()
        {
            ChangeSettings();
            RaiseChangeEvent();
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

            result[0] += ProjectWizardControl.CurrentLanguageID == 1033 ? "Project Type" + Environment.NewLine : "Projekt Typ" + Environment.NewLine;
            result[1] += SelectedProjectType(1033) + Environment.NewLine;

            result[0] += ProjectWizardControl.CurrentLanguageID == 1033 ? "Project Folder" : "Projekt Ordner";
            result[1] += SelectedProjectFolderType(1033);

            return result;
        }

        #endregion

        internal bool IsToolAddinProject
        {
            get
            {
                return checkBoxUseTools.Checked;
            }
        }

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

        internal string SelectedProjectType(int languageID)
        {
            if (radioButtonAutomationAddin.Checked)
                return languageID == 1033 ? "AutomationAddin" : "AutomationAddin";
            if (radioButtonWindowsForms.Checked)
                return languageID == 1033 ? "WindowsForms" : "WindowsForms";
            if (radioButtonConsole.Checked)
                return languageID == 1033 ? "Console" : "Konsole";
            if (radioButtonClassLibrary.Checked)
                return languageID == 1033 ? "ClassLibrary" : "Klassenbibliothek";
            throw new IndexOutOfRangeException("SelectedProjectType");
        }

        internal string SelectedProjectFolderType(int languageID)
        {
            if (radioButtonApplicationData.Checked)
                return languageID == 1033 ? "ApplicationData" : "ApplicationData";
            if (radioButtonDesktop.Checked)
                return languageID == 1033 ? "Desktop" : "Desktop";
            if (radioButtonUserFolder.Checked)
                return languageID == 1033 ? "User" : "Eigene Dateien";
            if (radioButtonVSProjectFolder.Checked)
                return languageID == 1033 ? "VSProject" : "Visual Studio Projektordner";                   
            if (radioButtonCustomFolder.Checked)
                return languageID == 1033 ? "Custom" : "Benutzerdefiniert";  
            throw new IndexOutOfRangeException("SelectedProjectFolderType");
        }

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
        
        internal string SelectedFolder
        {
            get
            {
                switch (SelectedProjectFolderType(1033))
                {
                    case "Custom":
                        return textBoxCustomFolder.Text;
                    default:
                        return "";
                }
            }
        }

        private void ChangeSettings()
        {
            _settings.FirstChild.SelectSingleNode("ProjectType").InnerText = SelectedProjectType(1033);
            _settings.FirstChild.SelectSingleNode("ProjectFolderType").InnerText = SelectedProjectFolderType(1033);
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

        private void radioButtonProjectType_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxUseTools.Enabled = radioButtonAutomationAddin.Checked;

            ChangeSettings();
            RaiseChangeEvent();
        }

        private void radioButtonProjectFolder_CheckedChanged(object sender, EventArgs e)
        {
            ChangeSettings();
            RaiseChangeEvent();

            buttonChooseFolder.Enabled = radioButtonCustomFolder.Checked;
            if (radioButtonCustomFolder.Checked && string.IsNullOrWhiteSpace(textBoxCustomFolder.Text))
                buttonChooseFolder_Click(buttonChooseFolder, new EventArgs());
        }

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            try
            {
                 FolderBrowserDialog dialog = new FolderBrowserDialog();
                if (DialogResult.OK == dialog.ShowDialog(this))
                    textBoxCustomFolder.Text = dialog.SelectedPath;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, 1033);
                errorForm.ShowDialog(this);
            }
        }

        private bool IsAdministrator()
        {
            WindowsIdentity myWindowsIdentity = WindowsIdentity.GetCurrent();
            WindowsPrincipal myWindowsPrincipal = new WindowsPrincipal(myWindowsIdentity);
            return myWindowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void linkLabelNSTOInfo_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                switch (ProjectWizardControl.CurrentLanguageID)
                {
                    case 1049:
                        System.Diagnostics.Process.Start("http://netoffice.codeplex.com/wikipage?title=Tools_RS");
                        break;
                    case 1031:
                        System.Diagnostics.Process.Start("http://netoffice.codeplex.com/wikipage?title=Tools_DE");
                        break;
                    default:
                        System.Diagnostics.Process.Start("http://netoffice.codeplex.com/wikipage?title=Tools_EN");
                        break;
                }
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, 1033);
            }
        }
    }
}
