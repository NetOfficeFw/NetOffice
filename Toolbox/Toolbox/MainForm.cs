using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Xml;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Reflection;

using NetOffice.DeveloperToolbox.ApplicationObserver;
using NetOffice.DeveloperToolbox.RegistryEditor;
using NetOffice.DeveloperToolbox.AddinGuard;
using NetOffice.DeveloperToolbox.OfficeCompatibility;
using NetOffice.DeveloperToolbox.OutlookSecurity;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// mainform in the application
    /// </summary>
    public partial class MainForm : Form
    {
        #region Construction

        /// <summary>
        /// Designtime Ctor
        /// </summary>
        public MainForm(): this(new string[0])
        {
            InitializeComponent();
        } 

        /// <summary>
        /// Runtime Ctor
        /// </summary>
        /// <param name="args">commandline argument array</param>
        public MainForm(string[] args)
        {
            try
            {
                InitializeComponent();
                
                ToolboxControls = new List<IToolboxControl>();
                ResizeControls();
                
                // setup about
                labelVersionText.Text = String.Format("Version {0}", AssemblyInfo.AssemblyVersion);
                labelVersionHint.Text = String.Format("Version {0}", AssemblyInfo.AssemblyVersion);
                labelCopyrightText.Text = AssemblyInfo.AssemblyCopyright;
                linkLabelCompany.Text = AssemblyInfo.AssemblyCompany;

                // load controls
                IntPtr dummyTabPageInsertDoesntWorkWithout = this.tabControlMain.Handle;
                OfficeCompatibilityControl newControl1 = new OfficeCompatibilityControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl1.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl1);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl1;
                imageListTabMain.Images.Add(newControl1.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl1.Dock = DockStyle.Fill;
                ToolboxControls.Add(newControl1);

                ApplicationObserverControl newControl2 = new ApplicationObserverControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl2.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl2);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl2;
                imageListTabMain.Images.Add(newControl2.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl2.Dock = DockStyle.Fill;
                ToolboxControls.Add(newControl2);

                RegistryEditorControl newControl3 = new RegistryEditorControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl3.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl3);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl3;
                imageListTabMain.Images.Add(newControl3.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl3.Dock = DockStyle.Fill;
                ToolboxControls.Add(newControl3);

                AddinGuardControl newControl4 = new AddinGuardControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl4.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl4);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl4;
                imageListTabMain.Images.Add(newControl4.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl4.Dock = DockStyle.Fill;
                ToolboxControls.Add(newControl4);

                OfficeUIControl newControl6 = new OfficeUIControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl6.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl6);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl6;
                imageListTabMain.Images.Add(newControl6.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl6.Dock = DockStyle.Fill;
                ToolboxControls.Add(newControl6);

                OutlookSecurityControl newControl5 = new OutlookSecurityControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl5.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl5);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl5;
                imageListTabMain.Images.Add(newControl5.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl5.Dock = DockStyle.Fill;
                ToolboxControls.Add(newControl5);

                ProjectWizardControl newControl7 = new ProjectWizardControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl7.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl7);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl7;
                imageListTabMain.Images.Add(newControl7.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl7.Dock = DockStyle.Fill;
                ToolboxControls.Add(newControl7);


                // load configuration
                LoadConfiguration();
                SetupTrayIcon(true);
                CommandLineArgs = args;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.Critical, CurrentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion
  
        #region Properties

        /// <summary>
        /// represents the tray icon in the right bottom area
        /// </summary>
        private NotifyIcon TrayIcon{get; set;}

        /// <summary>
        /// given commandline arguments from the application
        /// </summary>
        private string[] CommandLineArgs { get; set; }

        /// <summary>
        /// represents the LCID from the current ui language
        /// </summary>
        private int CurrentLanguageID { get; set; }

        /// <summary>
        /// represents the configuration file (old XmlDocument based because the first toolbox version was compiled in .NET2)
        /// </summary>
        private XmlDocument ConfigFile { get; set; }

        /// <summary>
        /// the currenty loaded toolbox controls
        /// </summary>
        private List<IToolboxControl> ToolboxControls { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// validate the version entry in the configuration file match with current assembly version
        /// </summary>
        /// <returns>true if match otherwise false</returns>
        private bool ValidateConfigFileVersion()
        {
            XmlAttribute versionAttribute = null;
            foreach (XmlAttribute item in ConfigFile.DocumentElement.Attributes)
            {
                if (item.Name == "Version")
                {
                    versionAttribute = item;
                    break;
                }
            }

            if (null == versionAttribute)
                return false;

            return (versionAttribute.Value == AssemblyInfo.AssemblyVersion);
        }

        /// <summary>
        /// Load the configuration file from LocalApplicationData folder or load the Ressources.DefaultConfiguration.xml(embedded ressource)
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DeveloperToolbox.Settings.xml");
                ConfigFile = new XmlDocument();
                if (File.Exists(filePath))
                {
                    ConfigFile.Load(filePath);
                    if (!ValidateConfigFileVersion())
                    {
                        File.Delete(filePath);
                        string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                        ConfigFile.Load(this.GetType().Assembly.GetManifestResourceStream(assemblyName + ".DefaultConfiguration.xml"));
                    }
                }
                else
                {
                    string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                    ConfigFile.Load(this.GetType().Assembly.GetManifestResourceStream(assemblyName + ".DefaultConfiguration.xml"));
                }

                XmlNode appNode = ConfigFile.SelectSingleNode("Application");
                checkBoxStartAppMinimized.Checked = Convert.ToBoolean(appNode.Attributes.Item(0).Value);
                checkBoxMinimizeToTray.Checked = Convert.ToBoolean(appNode.Attributes.Item(1).Value);
                checkBoxStartAppWithWindows.Checked = Convert.ToBoolean(appNode.Attributes.Item(2).Value);
                comboBoxLanguage.SelectedIndex = Convert.ToInt32((appNode.Attributes.Item(3).Value));

                foreach (IToolboxControl item in ToolboxControls)
                {
                    item.LoadConfiguration(ConfigFile.SelectSingleNode("Application/Controls/" + item.ControlName));
                    item.SetLanguage(IndexToLanguageID(comboBoxLanguage.SelectedIndex));
                }
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        /// <summary>
        /// Save the configuration to LocalApplicationData folder
        /// </summary>
        private void SaveConfiguration()
        {
            try
            {
                XmlNode appNode = ConfigFile.SelectSingleNode("Application");
                appNode.Attributes.Item(0).Value = checkBoxStartAppMinimized.Checked.ToString();
                appNode.Attributes.Item(1).Value = checkBoxMinimizeToTray.Checked.ToString();
                appNode.Attributes.Item(2).Value = checkBoxStartAppWithWindows.Checked.ToString();
                appNode.Attributes.Item(3).Value = comboBoxLanguage.SelectedIndex.ToString();

                foreach (IToolboxControl item in ToolboxControls)
                    item.SaveConfiguration(ConfigFile.SelectSingleNode("Application/Controls/" + item.ControlName));

                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DeveloperToolbox.Settings.xml");
                ConfigFile.Save(filePath);
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        /// <summary>
        /// Create autorun key in the registry or delete them
        /// </summary>
        private void SetupAutoRunEntry()
        {
            string runEntryKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
            string runEntryTitle = "DeveloperToolbox";

            if (checkBoxStartAppWithWindows.Checked)
            {
                RegistryKey runKey = Registry.CurrentUser.OpenSubKey(runEntryKey, true);
                object val = runKey.GetValue(runEntryTitle);
                if (val == null)
                    runKey.SetValue(runEntryTitle, this.GetType().Assembly.Location);
                runKey.Close();
            }
            else
            {
                RegistryKey runKey = Registry.CurrentUser.OpenSubKey(runEntryKey, true);
                object val = runKey.GetValue(runEntryTitle);
                if (val != null)
                    runKey.DeleteValue(runEntryTitle);
                runKey.Close();
            }
        }

        /// <summary>
        /// Create a TrayIcon(rightBottom corner)
        /// </summary>
        /// <param name="init">true means create the icon otherwise it means destroy the current tray icon</param>
        private void SetupTrayIcon(bool init)
        {
            if (true == init)
            {
                TrayIcon = new NotifyIcon();
                TrayIcon.Icon = this.Icon;
                TrayIcon.Text = "DeveloperToolbox";
                TrayIcon.Click += new EventHandler(TrayNotifyIcon_Click);
            }
            else
            {
                TrayIcon.Visible = false;
                TrayIcon.Dispose();
                TrayIcon = null;
            }
        }
      
        /// <summary>
        /// center the about page an title label
        /// </summary>
        private void ResizeControls()
        {
            colorLabelTitle.Top = (pictureBoxLogo.Top / 2) - (colorLabelTitle.Height / 2);
            colorLabelTitle.Left = (tabPageApplication.Width / 2) - (colorLabelTitle.Width / 2);

            panelAbout.Left = (tabPageAbout.Width / 2) - (panelAbout.Width / 2);
            panelAbout.Top = (tabPageAbout.Height / 2) - (panelAbout.Height / 2);
        }

        /// <summary>
        /// converts the language LCID parameter to a number and returns the language combobox index 
        /// </summary>
        /// <param name="language">language(must be a supported LCID)</param>
        /// <returns>combo box index</returns>
        private static int LanguageIDToIndex(string language)
        {
            int lang = Convert.ToInt32(language);
            if (lang == 1031)
                return 1;
            else
                return 0;
        }

        /// <summary>
        /// await a index from the language combobox and returns the corresonding LCID
        /// </summary>
        /// <param name="index">combo box index</param>
        /// <returns>coressponding LCID from index</returns>
        private static int IndexToLanguageID(int index)
        {
            if (0 == index)
                return 1033;
            else
                return 1031;
        }

        #endregion

        #region UI Trigger

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                foreach (var item in ToolboxControls)
                {
                    Control winControl = item as Control;
                    if (winControl.Visible)
                    {
                        item.KeyDown(e);
                        return;
                    }
                }
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        private void PictureBoxLogo_Click(object sender, EventArgs e)
        {
            try
            {
                tabControlMain.SelectedIndex = 1;
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        private void ComboBoxLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int languageID = IndexToLanguageID(comboBoxLanguage.SelectedIndex);
                CurrentLanguageID = languageID;
                Translator.TranslateControls(this, "Ressources.MessageTable.txt", languageID);
                foreach (IToolboxControl item in ToolboxControls)
                    item.SetLanguage(languageID);
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                SaveConfiguration();
                foreach (IToolboxControl item in ToolboxControls)
                    item.Dispose();

                SetupTrayIcon(false);
                SetupAutoRunEntry();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        private void TrayNotifyIcon_Click(object sender, EventArgs e)
        {
            try
            {
                this.Show();
                TrayIcon.Visible = false;
                this.WindowState = FormWindowState.Normal;
                this.Activate();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            try
            {
                if ((FormWindowState.Minimized == this.WindowState) && (true == checkBoxMinimizeToTray.Checked))
                {
                    TrayIcon.Visible = true;
                    this.Hide();
                }
                else
                {
                    TrayIcon.Visible = false;
                    ResizeControls();
                }

            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);              
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            try
            {
                bool minimize = false;

                if (true == checkBoxStartAppMinimized.Checked)
                    minimize = true;

                if (null != CommandLineArgs)
                {
                    foreach (string item in CommandLineArgs)
                        if (true == item.Trim().ToLower().Equals("-min", StringComparison.InvariantCultureIgnoreCase))
                            minimize = true;
                }

                if (minimize)
                    this.WindowState = FormWindowState.Minimized;
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);      
            }
        }

        private void TabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                TabPage currentTab = tabControlMain.TabPages[tabControlMain.SelectedIndex];
                IToolboxControl control = currentTab.Tag as IToolboxControl;
                if (null != control)
                    control.Activate();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);        
            }
        }

        private void LinkLabelAbout_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                System.Diagnostics.Process.Start(label.Text);
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        private void LinkLabelMain_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                System.Diagnostics.Process.Start(label.Text);
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);   
            }           
        }
       
        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                foreach (IToolboxControl item in ToolboxControls)
                    item.LoadComplete();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical, CurrentLanguageID);
            }
        }

        #endregion        
    }
}
