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
    public partial class MainForm : Form
    {
        #region Fields
        
        NotifyIcon  _notify;
        string[]    _args;
        XmlDocument _configFile;
        List<IToolboxControl> _controls = new List<IToolboxControl>();
        int _currentLanguageID;

        #endregion

        #region Construction

        public MainForm(): this(new string[0])
        {
             
        } 

        public MainForm(string[] args)
        {
            try
            {
                InitializeComponent();
                ResizeControls();

                // setup about
                labelVersionText.Text = String.Format("Version {0}", AssemblyVersion);
                labelVersionHint.Text = String.Format("Version {0}", AssemblyVersion);
                labelCopyrightText.Text = AssemblyCopyright;
                linkLabelCompany.Text = AssemblyCompany;

                // load controls
                IntPtr dummyTabPageInsertDoesntWorkWithout = this.tabControlMain.Handle;
                OfficeCompatibilityControl newControl1 = new OfficeCompatibilityControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl1.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl1);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl1;
                imageListTabMain.Images.Add(newControl1.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl1.Dock = DockStyle.Fill;
                _controls.Add(newControl1);

                ApplicationObserverControl newControl2 = new ApplicationObserverControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl2.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl2);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl2;
                imageListTabMain.Images.Add(newControl2.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl2.Dock = DockStyle.Fill;
                _controls.Add(newControl2);

                RegistryEditorControl newControl3 = new RegistryEditorControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl3.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl3);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl3;
                imageListTabMain.Images.Add(newControl3.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl3.Dock = DockStyle.Fill;
                _controls.Add(newControl3);

                AddinGuardControl newControl4 = new AddinGuardControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl4.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl4);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl4;
                imageListTabMain.Images.Add(newControl4.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl4.Dock = DockStyle.Fill;
                _controls.Add(newControl4);

                OfficeUIControl newControl6 = new OfficeUIControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl6.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl6);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl6;
                imageListTabMain.Images.Add(newControl6.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl6.Dock = DockStyle.Fill;
                _controls.Add(newControl6);

                OutlookSecurityControl newControl5 = new OutlookSecurityControl();
                tabControlMain.TabPages.Insert(tabControlMain.TabPages.Count - 1, newControl5.ControlCaption);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Controls.Add(newControl5);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].Tag = newControl5;
                imageListTabMain.Images.Add(newControl5.Icon);
                tabControlMain.TabPages[tabControlMain.TabPages.Count - 2].ImageIndex = imageListTabMain.Images.Count - 1;
                newControl5.Dock = DockStyle.Fill;
                _controls.Add(newControl5);

                // load configuration
                LoadConfiguration();
                SetupTrayIcon(true);
                _args = args;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.Critical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion

        #region Assemblyattributaccessoren

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }

        #endregion

        #region Config Methods

        private void LoadConfiguration()
        {
            try
            {
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DeveloperToolbox.Settings.xml");
                _configFile = new XmlDocument();
                if (File.Exists(filePath))
                    _configFile.Load(filePath);
                else
                {
                    string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                    _configFile.Load(this.GetType().Assembly.GetManifestResourceStream(assemblyName + ".DefaultConfiguration.xml"));
                }

                XmlNode appNode = _configFile.SelectSingleNode("Application");
                checkBoxStartAppMinimized.Checked = Convert.ToBoolean(appNode.Attributes.Item(0).Value);
                checkBoxMinimizeToTray.Checked = Convert.ToBoolean(appNode.Attributes.Item(1).Value);
                checkBoxStartAppWithWindows.Checked = Convert.ToBoolean(appNode.Attributes.Item(2).Value);
                comboBoxLanguage.SelectedIndex = Convert.ToInt32((appNode.Attributes.Item(3).Value));
                 
                foreach (IToolboxControl item in _controls)
                { 
                    item.LoadConfiguration(_configFile.SelectSingleNode("Application/Controls/" + item.ControlName));
                    item.SetLanguage(IndexToLanguageID(comboBoxLanguage.SelectedIndex));
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void SaveConfiguration()
        {
            try
            {
                XmlNode appNode = _configFile.SelectSingleNode("Application");
                appNode.Attributes.Item(0).Value = checkBoxStartAppMinimized.Checked.ToString();
                appNode.Attributes.Item(1).Value = checkBoxMinimizeToTray.Checked.ToString();
                appNode.Attributes.Item(2).Value = checkBoxStartAppWithWindows.Checked.ToString();
                appNode.Attributes.Item(3).Value = comboBoxLanguage.SelectedIndex.ToString();

                foreach (IToolboxControl item in _controls)
                    item.SaveConfiguration(_configFile.SelectSingleNode("Application/Controls/" + item.ControlName));

                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DeveloperToolbox.Settings.xml");
                _configFile.Save(filePath);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion

        #region Methods

        private static int LanguageIDToIndex(string language)
        {
            int lang = Convert.ToInt32(language);
            if (lang == 1031)
                return 1;
            else
                return 0;
        }

        private static int IndexToLanguageID(int index)
        {
            if (0 == index)
                return 1033;
            else
                return 1031;
        }

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

        private void SetupTrayIcon(bool init)
        {
            if (true == init)
            {
                _notify = new NotifyIcon();
                _notify.Icon = this.Icon;
                _notify.Text = "DeveloperToolbox";
                _notify.Click += new EventHandler(_notify_Click);
            }
            else
            {
                _notify.Visible = false;
                _notify.Dispose();
                _notify = null;
            }
        }

        #endregion

        #region Trigger

        private void comboBoxLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int languageID = IndexToLanguageID(comboBoxLanguage.SelectedIndex);
                _currentLanguageID = languageID;
                Translator.TranslateControls(this, "MessageTable.txt", languageID);
                foreach (IToolboxControl item in _controls)
                    item.SetLanguage(languageID);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                SaveConfiguration();
                foreach (IToolboxControl item in _controls)
                    item.Dispose();

                SetupTrayIcon(false);
                SetupAutoRunEntry();
            }
            catch (Exception exception)
            {
               ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
               errorForm.ShowDialog(this);
            }
        
        }
        
        void _notify_Click(object sender, EventArgs e)
        {
            try
            {
                this.Show();
                _notify.Visible = false;
                this.WindowState = FormWindowState.Normal;
                this.Activate();
            }
            catch (Exception exception)
            {
               ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
               errorForm.ShowDialog(this);
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            try
            {
                if ((FormWindowState.Minimized == this.WindowState) && (true == checkBoxMinimizeToTray.Checked))
                {
                    _notify.Visible = true;
                    this.Hide();
                }
                else
                {
                    _notify.Visible = false;
                    ResizeControls();
                }

            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);                
            }
        }

        private void ResizeControls()
        {
            colorLabelTitle.Top = (pictureBoxLogo.Top / 2) - (colorLabelTitle.Height / 2);
            colorLabelTitle.Left = (tabPageApplication.Width / 2) - (colorLabelTitle.Width / 2);

            panelAbout.Left = (tabPageAbout.Width / 2) - (panelAbout.Width / 2);
            panelAbout.Top = (tabPageAbout.Height / 2) - (panelAbout.Height / 2);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            try
            {
                bool minimize = false;

                if (true == checkBoxStartAppMinimized.Checked)
                    minimize = true;

                if (null != _args)
                {
                    foreach (string item in _args)
                        if (true == item.Trim().ToLower().Equals("-min", StringComparison.InvariantCultureIgnoreCase))
                            minimize = true;
                }

                if (minimize)
                    this.WindowState = FormWindowState.Minimized;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);             
            }
        }

        private void tabControlMain_SelectedIndexChanged(object sender, EventArgs e)
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
               ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
               errorForm.ShowDialog(this);             
            }
        }

        private void linkLabelAbout_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                System.Diagnostics.Process.Start(label.Text);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void linkLabelMain_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                System.Diagnostics.Process.Start(label.Text);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);     
            }           
        }
       
        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                foreach (IToolboxControl item in _controls)
                    item.LoadComplete();
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);     
            }
        }

        #endregion        
    }
}
