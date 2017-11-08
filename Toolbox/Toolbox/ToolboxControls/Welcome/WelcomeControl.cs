using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;

namespace NetOffice.DeveloperToolbox.ToolboxControls.Welcome
{
    /// <summary>
    /// The first title page of the application
    /// </summary>
    [RessourceTable("ToolboxControls.Welcome.Strings.txt")]
    public partial class WelcomeControl : UserControl, IToolboxControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public WelcomeControl()
        {
            InitializeComponent();
            labelVersionHint.Text = String.Format("Version {0}", AssemblyInfo.AssemblyVersion);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Corresponding tray icon in the bottom right area
        /// </summary>
        private NotifyIcon TrayIcon { get; set; }

        /// <summary>
        /// Last Right Click Link Label
        /// </summary>
        private LinkLabel LastClickedLinkLabel { get; set; }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public string ControlName
        {
            get { return "Welcome.WelcomeControl"; }
        }

        public string ControlCaption
        {
            get { return "Hallo"; }
        }

        public Image Icon
        {
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.Welcome.welcome.png"); }
        }

        public bool SupportsHelpContent
        {
            get 
            {
                return false;
            }
        }

        public bool SupportsInfoMessage
        {
            get
            {
                return false;
            }
        }

        public ToolboxControlMessageKind InfoMessageKind
        {
            get
            {
                return ToolboxControlMessageKind.Uncategorized;
            }
        }

        public string InfoMessage
        {
            get
            {
                return String.Empty;
            }
        }

        public void InitializeControl(IToolboxHost host)
        {            
            Host = host;
            Host.Minimized += new EventHandler(Host_Minimized);        
        }

        public void Activate(bool firstTime)
        {
            if (firstTime)
            {        
                controlForeColorAnimator1.Start(false);
                if (Program.IsAdmin)
                {
                    ApplicationStatePictureBox.Visible = false;
                    ApplicationStateLabel.Visible = false;
                    ApplicationStateLabel1.Visible = false;
                    ApplicationStateLabel2.Visible = false;
                    ApplicationStateLabel3.Visible = false;
                    linkLabelRunAsAdmin.Visible = false;
                }
            }
        }

        public void Deactivated()
        {

        }

        public Stream GetHelpText()
        {
            throw new NotSupportedException();
        }

        public void LoadComplete()
        {
            if (checkBoxStartAppMinimized.Checked)
                Host.MinimizeMainWindow(false);
        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
            XmlNode minimizeNode = configNode["StartMinimized"];
            XmlNode trayNode = configNode["MinimizeToTray"];
            XmlNode startupNode = configNode["RunAtStartup"];
            XmlNode languageNode = configNode["Language"];

            if (null != minimizeNode)
                checkBoxStartAppMinimized.Checked = Convert.ToBoolean(minimizeNode.InnerText);
            if(null != trayNode)
                checkBoxMinimizeToTray.Checked = Convert.ToBoolean(trayNode.InnerText);
            if(null != startupNode)
                checkBoxStartAppWithWindows.Checked = Convert.ToBoolean(startupNode.InnerText);
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
            XmlNode minimizeNode = configNode["StartMinimized"];
            XmlNode trayNode = configNode["MinimizeToTray"];
            XmlNode startupNode = configNode["RunAtStartup"];
            XmlNode languageNode = configNode["Language"];

            if (null == minimizeNode)
            { 
                minimizeNode = configNode.OwnerDocument.CreateNode(XmlNodeType.Element, "StartMinimized", null);
                configNode.AppendChild(minimizeNode);
            }

            if (null == trayNode)
            { 
                trayNode = configNode.OwnerDocument.CreateNode(XmlNodeType.Element, "MinimizeToTray", null);
                configNode.AppendChild(trayNode);
            }

            if (null == startupNode)
            { 
                startupNode = configNode.OwnerDocument.CreateNode(XmlNodeType.Element, "RunAtStartup", null);
                configNode.AppendChild(startupNode);
            }

            if (null == languageNode)
            { 
                languageNode = configNode.OwnerDocument.CreateNode(XmlNodeType.Element, "Language", null);
                configNode.AppendChild(languageNode);
            }

            minimizeNode.InnerText = checkBoxStartAppMinimized.Checked.ToString();
            trayNode.InnerText = checkBoxMinimizeToTray.Checked.ToString();
            startupNode.InnerText = checkBoxStartAppWithWindows.Checked.ToString();
  
            SetupAutoRunEntry();
        }

        public void SetLanguage(int id)
        {

        }

        public Stream GetHelpText(int lcid)
        {
            throw new NotImplementedException();
        }

        public new void KeyDown(KeyEventArgs e)
        {
         
        }

        public void Release()
        {
            Host.Minimized -= new EventHandler(Host_Minimized);
            SetupTrayIcon(false);
        }

        #endregion

        #region Methods

        private void SetupAutoRunEntry()
        {
            string runEntryKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
            string runEntryTitle = "NetOffice DeveloperToolbox";

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
                if (null != TrayIcon)
                {
                    TrayIcon.Click -= new EventHandler(TrayNotifyIcon_Click);
                    TrayIcon.Visible = false;
                    TrayIcon.Dispose();
                    TrayIcon = null;
                }

                TrayIcon = new NotifyIcon();
                TrayIcon.Icon = Host.Icon;
                TrayIcon.Text = Host.Caption;
                TrayIcon.Visible = true;
                TrayIcon.Click += new EventHandler(TrayNotifyIcon_Click);
                Host.MinimizeMainWindow(false);
            }
            else
            {
                if (null != TrayIcon)
                { 
                    TrayIcon.Click -= new EventHandler(TrayNotifyIcon_Click);
                    TrayIcon.Visible = false;
                    TrayIcon.Dispose();
                    TrayIcon = null;
                }
            }
        }

        private void ResizeControls()
        {
            pictureBoxHeader.Location = new Point((this.Width / 2) - (pictureBoxHeader.Width / 2), 13);
            labelVersionHint.Location = new Point((this.Width / 2) - (labelVersionHint.Width / 2)-5, pictureBoxLogo.Top + pictureBoxLogo.Height-40);
        }

        #endregion

        #region Trigger

        private void linkLabelRunAsAdmin_LinkClicked(object sender, LinkLabelLinkClickedEventArgs args)
        {
            try
            {
                Program.PerformSelfElevation(true);
                Application.Exit();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void LinkLabel_Clicked(object sender, EventArgs args)
        {
            try
            {
                MouseEventArgs mouseArgs = args as MouseEventArgs;
                if (null == mouseArgs)
                    return;

                if (mouseArgs.Button == MouseButtons.Left)
                {
                    LinkLabel label = sender as LinkLabel;
                    System.Diagnostics.Process.Start(label.Tag as string);
                }
                else if (mouseArgs.Button == MouseButtons.Right)
                {
                    LastClickedLinkLabel = sender as LinkLabel;
                    LinkContextMenu.Show(sender as Control, 0, 0);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void WelcomeControl_Resize(object sender, EventArgs e)
        {
            try
            {
                ResizeControls();
            }
            catch (Exception exception)
            {
                 Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void TrayNotifyIcon_Click(object sender, EventArgs e)
        {
            try
            {
                TrayIcon.Visible = false;
                Host.ShowMainWindow();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void Host_Minimized(object sender, EventArgs e)
        {
            try
            {
                if (checkBoxMinimizeToTray.Checked)
                    SetupTrayIcon(true);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void pictureBoxLogo_Click(object sender, EventArgs e)
        {
            try
            {
                Host.SwitchTo("ProjectWizard.ProjectWizardControl");
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void LinkContextMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (null != LastClickedLinkLabel)
            {
                Clipboard.SetText(LastClickedLinkLabel.Tag as string);
            }
        }

        #endregion
    }
}