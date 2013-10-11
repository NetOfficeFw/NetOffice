using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Application Main Form
    /// </summary>
    public partial class FormMain : Form, IApplicationHost
    {
        #region Fields

        /// <summary>
        /// of course, we need a lock instance for ui updates(async notification calls)
        /// </summary>
        private object _lockUI = new object();
        
        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public FormMain()
        {
            InitializeComponent();
            EnumerateScreens();
            ConsoleUI.Host = this;
            ChannelUI.Host = this;
            toolStripStatusLabelAttention.Visible = IsRunningWithAdminPrivileges;
            CustomConsoleList = new CustomConsoleCollection();
            LoadSettings();
            ChannelUI.UpdateViewOptions(ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Reports the console is currently empty
        /// </summary>
        private bool ConsoleIsEmtpy
        {
            get
            {
                if (ConsoleUI.Items.Plain.Count > 1 || ChannelUI.Items.Count > 1)
                    return false;

                if (ConsoleUI.Items.Plain.Count < 1 && ChannelUI.Items.Count < 1)
                    foreach (var item in CustomConsoleList)
                        if (item.Items.Plain.Count > 0)
                            return false;

                return true;
            }
        }

        /// <summary>
        /// Handle the application tray icon
        /// </summary>
        private TrayIconManager TrayHandler { get; set; }

        /// <summary>
        /// Handle incoming messages
        /// </summary>
        private INotificationProvider UpdateHandler { get; set; }
   
        /// <summary>
        /// Main Console
        /// </summary>
        private ConsoleViewControl ConsoleUI { get { return consoleViewMain; } }

        /// <summary>
        /// Channels
        /// </summary>
        private ChannelViewControl ChannelUI { get { return channelViewMain; } }

        /// <summary>
        /// Custom Console Views
        /// </summary>
        private CustomConsoleCollection CustomConsoleList { get; set; }

        /// <summary>
        /// Current visible control or null(about page is visible9
        /// </summary>
        private IApplicationControl CurrentControl
        {
            get
            {
                return TabControlMain.SelectedTab.Controls[0] as IApplicationControl;
            }
        }

        /// <summary>
        /// Get info the application is running with elevated rights
        /// </summary>
        private bool IsRunningWithAdminPrivileges
        {
            get
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Clear all view content
        /// </summary>
        private void ClearAllDisplayContent()
        {
            lock (_lockUI)
            {
                ConsoleUI.Clear();
                ChannelUI.Clear();
                foreach (var item in CustomConsoleList)
                    item.Clear();
                foreach (TabPage item in TabControlMain.TabPages)
                {
                    IApplicationControl appControl = item.Controls[0] as IApplicationControl;
                    if(null != appControl && !String.IsNullOrWhiteSpace(appControl.ControlName))
                        item.Text = appControl.ControlName;  
                }
                TabControlMain.TabPages[0].Text = "Console";
                TabControlMain.TabPages[1].Text = "Channels";
            }
        }

        /// <summary>
        /// Show all screens in combo box and preselect primary screen
        /// </summary>
        private void EnumerateScreens()
        {
            comboBoxAvailableScreens.DataSource = Screen.AllScreens;
            if (Screen.AllScreens.Length > 0)
            {
                for (int i = 0; i < Screen.AllScreens.Length; i++)
                {
                    if (Screen.AllScreens[i].Primary)
                    {
                        comboBoxAvailableScreens.SelectedIndex = i;
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Get settings Screen or default screen
        /// </summary>
        /// <returns></returns>
        private Screen GetSettingsScreen()
        {
            string screen = Properties.Settings.Default.WindowMonitor;
            Screen defaultScreen = null;
            for (int i = 0; i < Screen.AllScreens.Length; i++)
                if (Screen.AllScreens[i].Primary)
                {
                    defaultScreen = Screen.AllScreens[i];
                    break;
                }

            if (!String.IsNullOrWhiteSpace(screen))
                foreach (var item in Screen.AllScreens)
                    if (screen == item.DeviceName)
                        return item;

            return defaultScreen;
        }

        /// <summary>
        /// Get current selected window start location
        /// </summary>
        /// <returns></returns>
        private WindowCustomStartLocation GetCurrentSelectedWindowStartLocation()
        {
            RadioButton checkedRadioButton = null;
            foreach (Control item in panelSettings1.Controls)
            {
                RadioButton radioButton = item as RadioButton;
                if(null != radioButton && radioButton.Checked)
                {
                    checkedRadioButton = radioButton;
                    break;
                }
            }

            if (null == checkedRadioButton)
                return WindowCustomStartLocation.TopRight;

            switch (checkedRadioButton.Name)
            {
                case "radioButtonTopLeft":
                    return WindowCustomStartLocation.TopLeft;
                case "radioButtonTopRight":
                    return WindowCustomStartLocation.TopRight;
                case "radioButtonBottomLeft":
                    return WindowCustomStartLocation.BottomLeft;
                case "radioButtonBottomRight":
                    return WindowCustomStartLocation.BottomRight;
                case "radioButtonCenter":
                    return WindowCustomStartLocation.Center;
                case "radioButtonLastPosition":
                    return WindowCustomStartLocation.LastPosition;
                case "radioButtonMaximized":
                    return WindowCustomStartLocation.Maximized;
                default:
                    return WindowCustomStartLocation.TopRight;
            }
        }

        /// <summary>
        /// Load application settings
        /// </summary>
        private void LoadSettings()
        {
            Screen screen = GetSettingsScreen();
            
            if (null != screen)
            {
                WindowCustomStartLocation location = (WindowCustomStartLocation)Properties.Settings.Default.WindowStartLocation;
                switch (location)
                {
                    case WindowCustomStartLocation.TopLeft:
                        this.Top = screen.WorkingArea.Top;
                        this.Left = screen.WorkingArea.Left;
                        radioButtonTopLeft.Checked = true;
                        break;
                    case WindowCustomStartLocation.TopRight:
                        this.Top = screen.WorkingArea.Top;
                        this.Left = screen.WorkingArea.Size.Width - this.Width;
                        radioButtonTopRight.Checked = true;
                        break;
                    case WindowCustomStartLocation.BottomLeft:
                        this.Top = screen.WorkingArea.Size.Height - this.Height;
                        this.Left = screen.WorkingArea.Left;
                        radioButtonBottomLeft.Checked = true;
                        break;
                    case WindowCustomStartLocation.BottomRight:
                        this.Top = screen.WorkingArea.Size.Height - this.Height;
                        this.Left = screen.WorkingArea.Size.Width - this.Width;
                        radioButtonBottomRight.Checked = true;
                        break;
                    case WindowCustomStartLocation.Center:
                        this.Top = (screen.WorkingArea.Size.Height / 2) - (this.Height / 2);
                        this.Left = (screen.WorkingArea.Size.Width/ 2) - (this.Width / 2);
                        radioButtonCenter.Checked = true;
                        break;
                    case WindowCustomStartLocation.LastPosition:
                        this.Top = Properties.Settings.Default.LastPositionY;
                        this.Left = Properties.Settings.Default.LastPositionX;
                        radioButtonLastPosition.Checked = true;
                        break;
                    case WindowCustomStartLocation.Maximized:
                        this.Top = screen.WorkingArea.Top;
                        this.Left = screen.WorkingArea.Left;
                        this.WindowState = FormWindowState.Maximized;
                        radioButtonMaximized.Checked = true;
                        break;
                    default:
                        break;
                }

                checkBoxStartInTray.Checked = Properties.Settings.Default.WindowStartInTray;
                ToolStripMenuItemShowTime.Checked = Properties.Settings.Default.ShowTime;
                ToolStripMenuItemShowMachine.Checked = Properties.Settings.Default.ShowMachine;
                ToolStripMenuItemShowAppDomain.Checked = Properties.Settings.Default.ShowAppDomain;
                ToolStripMenuItemAlwaysOnTop.Checked = Properties.Settings.Default.WindowAlwaysOnTop;
                ToolStripMenuItemEnabled.Checked = Properties.Settings.Default.Enabled;

                if (Properties.Settings.Default.WindowStartInTray)
                    this.WindowState = FormWindowState.Minimized;
            }
        }

        /// <summary>
        /// Save application settings
        /// </summary>
        private void SaveSettings()
        {
            Screen screen = comboBoxAvailableScreens.SelectedItem as Screen;
            if (null == screen)
                screen = GetSettingsScreen();
            if (null == screen) return;

            Properties.Settings.Default.WindowMonitor = screen.DeviceName;
            Properties.Settings.Default.WindowStartLocation = Convert.ToInt32(GetCurrentSelectedWindowStartLocation());
            Properties.Settings.Default.WindowStartInTray = checkBoxStartInTray.Checked;

            Properties.Settings.Default.Enabled = ToolStripMenuItemEnabled.Checked;
            Properties.Settings.Default.WindowAlwaysOnTop = ToolStripMenuItemAlwaysOnTop.Checked;
            Properties.Settings.Default.ShowTime = ToolStripMenuItemShowTime.Checked;
            Properties.Settings.Default.ShowMachine = ToolStripMenuItemShowMachine.Checked;
            Properties.Settings.Default.ShowAppDomain = ToolStripMenuItemShowAppDomain.Checked;

            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Create a summary display string of all content
        /// </summary>
        /// <returns>summary</returns>
        private string CreateSummaryString()
        {
            StringBuilder sb = new StringBuilder(2404);
            if (ConsoleUI.Items.Plain.Count > 0)
                sb.Append("--Console Main--" + Environment.NewLine + ConsoleUI.Items.CreateText(ConsoleUI.ViewStyle, true, true, true));
            foreach (var item in CustomConsoleList)
                sb.Append("Console " + item.ControlName + Environment.NewLine + item.Items.CreateText(item.ViewStyle, true, true, true));

            sb.Append("--Channels--" + Environment.NewLine + ChannelUI.Items.CreateText(true, true, true));

            return sb.ToString();
        }

        /// <summary>
        /// Copy all content to clipboard
        /// </summary>
        private void CopyDisplayContentToClipboard()
        {
            lock (_lockUI)
            {
                Clipboard.SetText(CreateSummaryString(), TextDataFormat.Text);
            }
        }

        private void SaveDisplayContentToTextFile()
        {
            try
            {
                if (ConsoleIsEmtpy)
                {
                    MessageBox.Show(this, "The monitor is currently empty." + Environment.NewLine + "No need for save!", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "Plain/Text|*.txt";
                dialog.OverwritePrompt = true;
                DialogResult dr = dialog.ShowDialog(this);

                lock (this)
                {
                    if (dr == System.Windows.Forms.DialogResult.OK)
                    {
                        if (System.IO.File.Exists(dialog.FileName))
                            System.IO.File.Delete(dialog.FileName);
                        System.IO.File.AppendAllText(dialog.FileName, CreateSummaryString());
                    }
                }
            }
            catch (Exception exception)
            {
                string message = String.Format("An error is occured.{0}{0}{1}", Environment.NewLine, exception);
                MessageBox.Show(this, message, "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }            
        }

        /// <summary>
        /// Display incoming console message
        /// </summary>
        /// <param name="notifyTime"></param>
        /// <param name="appDomainFriendlyName"></param>
        /// <param name="name"></param>
        /// <param name="message"></param>
        private string ShowConsoleMessage(string notifyTime, string consoleName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain)
        {
            if (ConsoleUI.InvokeRequired)
                return ConsoleUI.Invoke(new UpdateMonitorInvoker(ShowConsoleMessage), new object[] { notifyTime, consoleName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain }) as string;
                 
            lock (_lockUI)
            {
                TrayHandler.ShowUpdate();
                if (String.IsNullOrWhiteSpace(consoleName))
                {
                    string resultKey = ConsoleUI.AddNewMessage(notifyTime, consoleName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain);
                    if (TabControlMain.SelectedIndex != 0)
                        TabControlMain.TabPages[0].Text = "*Console";
                    return resultKey;
                }
                else
                {
                    foreach (var item in CustomConsoleList)
                    {
                        if (item.ControlName.Trim().ToLower() == consoleName.Trim().ToLower())
                        {
                            string newEntryID = item.AddNewMessage(notifyTime, consoleName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain);

                            TabPage page = item.Parent as TabPage;
                            if (TabControlMain.SelectedTab != page)
                                page.Text = "*" + item.ControlName;
                            
                            return newEntryID;
                        }
                    }
                    return AddNewConsoleAndShowMessage(notifyTime, consoleName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain);
                }
            }
        }

        /// <summary>
        /// Display incoming channel message
        /// </summary>
        /// <param name="notifyTime"></param>
        /// <param name="appDomainFriendlyName"></param>
        /// <param name="name"></param>
        /// <param name="message"></param>
        private string ShowChannelMessage(string notifyTime, string channelName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain)
        {
            if (ChannelUI.InvokeRequired)
                return ChannelUI.Invoke(new UpdateMonitorInvoker(ShowChannelMessage), new object[] { notifyTime, channelName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain }) as string;

            lock (_lockUI)
            {
                TrayHandler.ShowUpdate();
                if (CurrentControl != ChannelUI)
                    TabControlMain.TabPages[1].Text = "*Channels";
                return ChannelUI.AddNewMessage(notifyTime, channelName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain);
            }
        }
        
        /// <summary>
        /// Creates a new console tabpage and display console message. Called in invoke/locked state from ShowConsoleMessage
        /// </summary>
        /// <param name="notifyTime"></param>
        /// <param name="appDomainFriendlyName"></param>
        /// <param name="name"></param>
        /// <param name="message"></param>
        private string AddNewConsoleAndShowMessage(string notifyTime, string consoleName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain)
        {
            if (TabControlMain.TabPages.Count > 32)
                return null;
            string validatedName = consoleName.Length > 10 ? consoleName.Substring(0, 9) : consoleName;
            ConsoleViewControl newConsole = new ConsoleViewControl(this, consoleName);
            newConsole.ShowCloseButton = true;            
            TabPage newConsolePage = new TabPage("*" + validatedName);
            TabControlMain.TabPages.Add(newConsolePage);
            newConsolePage.ImageIndex = 0;
            newConsolePage.Controls.Add(newConsole);
            newConsole.Dock = DockStyle.Fill;
            CustomConsoleList.Add(newConsole);
            string newEntryID = newConsole.AddNewMessage(notifyTime, consoleName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain);          
            newConsole.CloseClick += new EventHandler(CustomConsole_CloseClick);
            return newEntryID;
        }

        #endregion

        #region IApplicationHost
        
        public bool ShowTime
        {
            get { return ToolStripMenuItemShowTime.Checked; }
        }

        public bool ShowMachine
        {
            get { return ToolStripMenuItemShowMachine.Checked; }
        }

        public bool ShowAppDomain
        {
            get { return ToolStripMenuItemShowAppDomain.Checked; }
        }

        public bool IsCurrentlyVisible(IApplicationControl control)
        {
            return control == CurrentControl;
        }

        #endregion

        #region Trigger

        private void CustomConsole_CloseClick(object sender, EventArgs e)
        {
            ConsoleViewControl console = sender as ConsoleViewControl;
            lock (_lockUI)
            {
                console.CloseClick -= new EventHandler(CustomConsole_CloseClick);
                CustomConsoleList.Remove(console);
                TabPage consolePage = console.Parent as TabPage;
                TabControlMain.TabPages.Remove(consolePage);
            }
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            TrayHandler = new TrayIconManager(this, ContextMenuStripTray);
            UpdateHandler = new NamedPipes.PipeServer();
            UpdateHandler.ConsoleNotification += new UpdateMonitorInvoker(UpdateHandler_ConsoleNotification);
            UpdateHandler.ChannelNotification += new UpdateMonitorInvoker(UpdateHandler_ChannelNotification);
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
            UpdateHandler.Dispose();
            TrayHandler.Dispose();
        }

        private void TimerMain_Tick(object sender, EventArgs e)
        {
            ToolStripLabelTime.Text = DateTime.Now.ToLongTimeString();
        }
      
        private void ToolStripMenuItemEnabled_CheckedChanged(object sender, EventArgs e)
        {
            if (ToolStripMenuItemEnabled.Checked)
            {
                UpdateHandler.Start();
                Text = "DebugConsole Monitor - Enabled";
            }
            else
            { 
                UpdateHandler.Stop();
                Text = "DebugConsole Monitor - Disabled";
            }
        }

        private void ToolStripMenuItemAlwaysOnTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = ToolStripMenuItemAlwaysOnTop.Checked;
        }

        private void ToolStripMenuItemShowAppDomain_CheckedChanged(object sender, EventArgs e)
        {
            lock (_lockUI)
            {
                IApplicationControl control = CurrentControl;
                if (null != control)
                    control.UpdateDisplayContent(ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
            }
        }

        private void ToolStripMenuItemShowMachine_CheckedChanged(object sender, EventArgs e)
        {
            lock (_lockUI)
            {
                IApplicationControl control = CurrentControl;
                if (null != control)
                    control.UpdateDisplayContent(ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
            }
        }

        private void ToolStripMenuItemShowTime_CheckedChanged(object sender, EventArgs e)
        {
            lock (_lockUI)
            {
                IApplicationControl control = CurrentControl;
                if (null != control)
                    control.UpdateDisplayContent(ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
            }
        }

        private void ToolStripStatusLabelAttention_Click(object sender, EventArgs e)
        {
            FormAdminMessage.ShowAdminMessage(this);
        }

        private void ToolStripMenuItemClearConsole_Click(object sender, EventArgs e)
        {
            ClearAllDisplayContent();
        }

        private void ToolStripMenuItemSaveContent_Click(object sender, EventArgs e)
        {
            SaveDisplayContentToTextFile();
        }

        private void ToolStripMenuItemCopyContent_Click(object sender, EventArgs e)
        {
            CopyDisplayContentToClipboard();
        }

        private string UpdateHandler_ConsoleNotification(string notifyTime, string consoleName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain)
        {
            return ShowConsoleMessage(notifyTime, consoleName, machineName, appDomainFriendlyName, parentEntryID,  message, ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
        }

        private string UpdateHandler_ChannelNotification(string notifyTime, string channelName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain)
        {
            return ShowChannelMessage(notifyTime, channelName, machineName, appDomainFriendlyName, parentEntryID, message, ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
        }

        private void ToolStripMenuItemExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void linkLabelSuggestions1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(linkLabelSuggestions1.Text);
            }
            catch
            {
                ;
            }
        }
        private void linkLabelInfo_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(linkLabelInfo.Tag as string);
            }
            catch
            {
                ;
            }
        }

        private void FormMain_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
                this.Hide();
        }

        private void TabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (_lockUI)
            {
                IApplicationControl control = CurrentControl;
                if (null != control && !string.IsNullOrWhiteSpace(control.ControlName))
                {
                    control.UpdateDisplayContent(ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
                    TabControlMain.SelectedTab.Text = control.ControlName;
                }
                else 
                {
                    switch (TabControlMain.SelectedIndex)
                    {
                        case 0:
                            TabControlMain.TabPages[0].Text = "Console";
                            ConsoleUI.UpdateDisplayContent(ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
                            break;
                        case 1:
                            TabControlMain.TabPages[1].Text = "Channels";
                            ChannelUI.UpdateDisplayContent(ToolStripMenuItemShowTime.Checked, ToolStripMenuItemShowMachine.Checked, ToolStripMenuItemShowAppDomain.Checked);
                            break;
                    }
                }
            }
        }

        #endregion
    }
}
