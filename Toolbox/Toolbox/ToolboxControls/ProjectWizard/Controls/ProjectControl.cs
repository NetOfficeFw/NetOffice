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
    //step 1
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.ProjectControl.txt")]
    public partial class ProjectControl : UserControl, IWizardControl, ILocalizationDesign
    {
        private XmlDocument _settings;

        public ProjectControl()
        {
            InitializeComponent(); 
            CreateSettingsDocument();
            if (!IsAdministrator())
                labelNoAdminHint.Visible = true;
            else
                labelNoAdminHint.Visible = false;
        }

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
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
                    return "Project Typ";
                else
                    return "Project Type";
            }
        }

        public string Description
        {
            get
            {
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
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
            Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
            if (null != language)
            {
                var component = language.Components["Project Wizard - Project"];
                Translation.Translator.TranslateControls(this, component.ControlRessources);
            }
            else
            {
                Translation.Translator.TranslateControls(this, "ToolboxControls.ProjectWizard.Controls.ProjectControl.txt", Forms.MainForm.Singleton.CurrentLanguageID);
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

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {
            labelNoAdminHint.Visible = true;
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

        internal ProjectType SelectedProjectType()
        {
            
            if (radioButtonAutomationAddin.Checked)
            {
                if(checkBoxUseTools.Checked)
                    return ProjectType.NetOfficeAddin;
                else
                    return ProjectType.NetOfficeAddin;
            }
            if (radioButtonWindowsForms.Checked)
                return ProjectType.WindowsForms;
            if (radioButtonConsole.Checked)
                return  ProjectType.Console;
            if (radioButtonClassLibrary.Checked)
                return  ProjectType.ClassLibrary;
            throw new IndexOutOfRangeException("SelectedProjectType");
        }

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
                switch (SelectedProjectFolderType())
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

        private void radioButtonProjectType_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxUseTools.Enabled = radioButtonAutomationAddin.Checked;
            checkBoxUseTools.Checked = radioButtonAutomationAddin.Checked;
            ChangeSettings();
            RaiseChangeEvent();
        }

        private void radioButtonProjectFolder_CheckedChanged(object sender, EventArgs e)
        {
            ChangeSettings();
            RaiseChangeEvent();

            RadioButton button = sender as RadioButton;
            buttonChooseFolder.Enabled = radioButtonCustomFolder.Checked;            
            if (radioButtonCustomFolder.Checked && string.IsNullOrWhiteSpace(textBoxCustomFolder.Text) & button == radioButtonCustomFolder)
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
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, 1033);
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
                switch (Forms.MainForm.Singleton.CurrentLanguageID)
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical, 1033);
            }
        }
    }
}
