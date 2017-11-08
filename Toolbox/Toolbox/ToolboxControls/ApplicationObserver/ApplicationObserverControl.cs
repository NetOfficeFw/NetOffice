using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using System.Diagnostics;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ApplicationObserver
{
    /// <summary>
    /// Observe MS processes and kill easy
    /// </summary>
    [RessourceTable("ToolboxControls.ApplicationObserver.Strings.txt")]
    public partial class ApplicationObserverControl : UserControl, IToolboxControl
    {
        #region Fields

        private OfficeApplicationObserver _applicationObserver;
     
        #endregion

        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ApplicationObserverControl()
        {
            try
            {
                InitializeComponent();
                if (!Program.IsDesign)
                {
                    CreateHandle();
                    _applicationObserver = new OfficeApplicationObserver(listViewApps);
                    textBoxHotKey.Text = _applicationObserver.HotKey.ToString();
                    _applicationObserver.InstanceRunningCountChanged += new EventHandler(ApplicationObserver_InstanceRunningCountChanged);
                    _applicationObserver.AllProcessesChanged += new EventHandler(ApplicationObserver_AllProcessesChanged);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
        
        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public void InitializeControl(IToolboxHost host)
        {
            Host = host;
        }

        public new void KeyDown(KeyEventArgs e)
        { 
        }

        public string ControlName
        {
            get 
            {
                return "ApplicationObserver.ApplicationObserverControl";
            }
        }

        public string ControlCaption
        {
            get { return "Observer"; }
        }

        public System.ComponentModel.IContainer Components
        {
            get
            {
                return components;
            }
        }

        public Image Icon
        {
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.ApplicationObserver.Icon.png"); }
        }

        public bool SupportsHelpContent
        {
            get
            {
                return true;
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

        public void Activate(bool firstTime)
        {

        }

        public void Deactivated()
        {

        }

        public void LoadComplete()
        {
           
        }

        public void LoadConfiguration(XmlNode configNode)
        {
            if (configNode.ChildNodes.Count == 0)
                configNode.InnerXml = Ressources.RessourceUtils.ReadString("ToolboxControls.ApplicationObserver.IconsAndConfig.DefaultConfiguration.txt");
           
            string val = "";

            val = configNode.SelectSingleNode("Excel").Attributes[0].Value;
            listViewApps.Items[0].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Winword").Attributes[0].Value;
            listViewApps.Items[1].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Outlook").Attributes[0].Value;
            listViewApps.Items[2].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("PowerPnt").Attributes[0].Value;
            listViewApps.Items[3].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("MsAccess").Attributes[0].Value;
            listViewApps.Items[4].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Project").Attributes[0].Value;
            listViewApps.Items[5].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Visio").Attributes[0].Value;
            listViewApps.Items[6].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Tray").Attributes[0].Value;
            checkBoxAppsTray.Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("HotKey").Attributes[1].Value;
            textBoxHotKey.Text = ((Keys)Convert.ToInt32(val)).ToString();
            _applicationObserver.HotKey = (Keys)Convert.ToInt32(val);

            val = configNode.SelectSingleNode("HotKey").Attributes[0].Value;
            checkBoxAppKill.Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("HotKey").Attributes[2].Value;
            checkBoxShowQuestion.Checked = Convert.ToBoolean(val);
        }

        public void SaveConfiguration(XmlNode configNode)
        {
            if (configNode.ChildNodes.Count == 0)
                configNode.InnerXml = Ressources.RessourceUtils.ReadString("ToolboxControls.ApplicationObserver.IconsAndConfig.DefaultConfiguration.txt");

            configNode.SelectSingleNode("Excel").Attributes[0].Value = listViewApps.Items[0].Checked.ToString();
            configNode.SelectSingleNode("Winword").Attributes[0].Value = listViewApps.Items[1].Checked.ToString();
            configNode.SelectSingleNode("Outlook").Attributes[0].Value = listViewApps.Items[2].Checked.ToString();
            configNode.SelectSingleNode("PowerPnt").Attributes[0].Value = listViewApps.Items[3].Checked.ToString();
            configNode.SelectSingleNode("MsAccess").Attributes[0].Value = listViewApps.Items[4].Checked.ToString();
            configNode.SelectSingleNode("Project").Attributes[0].Value = listViewApps.Items[5].Checked.ToString();
            configNode.SelectSingleNode("Visio").Attributes[0].Value = listViewApps.Items[6].Checked.ToString();
            configNode.SelectSingleNode("Tray").Attributes[0].Value = checkBoxAppsTray.Checked.ToString();
            configNode.SelectSingleNode("HotKey").Attributes[0].Value = checkBoxAppKill.Checked.ToString();
            configNode.SelectSingleNode("HotKey").Attributes[1].Value = ((int)_applicationObserver.HotKey).ToString();
            configNode.SelectSingleNode("HotKey").Attributes[2].Value = checkBoxShowQuestion.Checked.ToString();
        }


        public Stream GetHelpText()
        {
                return Ressources.RessourceUtils.ReadStream("ToolboxControls.ApplicationObserver.Info1033.rtf");
        }
        
        public void Release()
        {
            if ((null != _applicationObserver) && (!Program.IsDesign))
            {
                _applicationObserver.Dispose();
                _applicationObserver = null;
            }
        }

        #endregion

        #region Methods

        private static int GetProcessImageIndex(string processName)
        {
            processName = processName.Trim().ToUpper();
            switch (processName)
            {
                case "EXCEL":
                    return 1;
                case "WINWORD":
                    return 2;
                case "OUTLOOK":
                    return 3;
                case "POWERPNT":
                    return 4;
                case "MSACCESS":
                    return 5;
                case "WINPROJ":
                    return 6;
                case "VISIO":
                    return 6;
                default:
                    return 0;
            }

        }

        #endregion

        #region Trigger

        private void ApplicationObserver_AllProcessesChanged(object sender, EventArgs e)
        {
            try
            {
                listViewProcess.Items.Clear();
                Process[] process = sender as Process[];
                foreach (Process item in process)
                {
                    ListViewItem vieItem = listViewProcess.Items.Add("");
                    vieItem.SubItems.Add(item.Id.ToString());
                    vieItem.SubItems.Add(item.ProcessName);
                    vieItem.ImageIndex = GetProcessImageIndex(item.ProcessName);
                }

                int i = 0;
                foreach (ListViewItem item in listViewProcess.Items)
                {
                    Color color = i % 2 != 0 ? Color.White : Color.LightGray;
                    item.BackColor = color;
                    foreach (ListViewItem.ListViewSubItem subItem in item.SubItems)
                        subItem.BackColor = color;
                    i++;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void ApplicationObserver_InstanceRunningCountChanged(object sender, EventArgs e)
        {
            try
            {
                foreach (ListViewItem item in listViewApps.Items)
                {
                    if (item.SubItems[1].Text.Length > 0)
                    {
                        int number = Convert.ToInt32(item.SubItems[1].Text);
                        if ((number > 0) && (item.Checked))
                        {
                            buttonKillApps.Enabled = true;
                            return;
                        }
                    }
                }
                buttonKillApps.Enabled = false;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void listViewApps_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
                string appName = e.Item.Text;

                if (((e.Item.SubItems[1].Text != "0") && (!buttonKillApps.Enabled)))
                    buttonKillApps.Enabled = true;

                switch (appName)
                {
                    case "Excel":
                        _applicationObserver.Excel = e.Item.Checked;
                        break;
                    case "Winword":
                        _applicationObserver.Word = e.Item.Checked;
                        break;
                    case "Outlook":
                        _applicationObserver.Outlook = e.Item.Checked;
                        break;
                    case "PowerPnt":
                        _applicationObserver.PowerPoint = e.Item.Checked;
                        break;
                    case "MsAccess":
                        _applicationObserver.Access = e.Item.Checked;
                        break;
                    case "WinProj":
                        _applicationObserver.Project = e.Item.Checked;
                        break;
                    case "Visio":
                        _applicationObserver.Visio = e.Item.Checked;
                        break;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void checkBoxAppsTray_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                _applicationObserver.TrayIcon = checkBoxAppsTray.Checked;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void checkBoxAppKill_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                _applicationObserver.HotKeyEnabled = checkBoxAppKill.Checked;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void textBoxHotKey_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (true == checkBoxAppKill.Checked)
                    checkBoxAppKill.Checked = false;

                string CtrlKeys = "";
                if (e.Control)
                    CtrlKeys += "Ctrl ";
                if (e.Alt)
                    CtrlKeys += "Alt ";

                if (e.KeyCode == (Keys.LButton | Keys.ShiftKey))
                    textBoxHotKey.Text = CtrlKeys;
                else if (e.KeyCode == Keys.Menu)
                    textBoxHotKey.Text = CtrlKeys;
                else
                    textBoxHotKey.Text = CtrlKeys + e.KeyCode.ToString();

                if ((e.Control) && (e.Alt))
                    _applicationObserver.HotKey = e.KeyCode | Keys.Control | Keys.Alt;
                else if (e.Control)
                    _applicationObserver.HotKey = e.KeyCode | Keys.Control;
                else if (e.Alt)
                    _applicationObserver.HotKey = e.KeyCode | Keys.Alt;
                else
                    _applicationObserver.HotKey = e.KeyCode;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void buttonKillApps_Click(object sender, EventArgs e)
        {
            try
            {
                _applicationObserver.KillProcesses();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void labelKillQuestion_TextChanged(object sender, EventArgs e)
        {
            try
            {
                _applicationObserver.KillQuestion = labelKillQuestion.Text;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void checkBoxShowQuestion_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                _applicationObserver.ShowQuestionBeforeKill = checkBoxShowQuestion.Checked;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void listViewProcess_Resize(object sender, EventArgs e)
        {
            try
            {
                listViewProcess.Columns[2].Width = (listViewProcess.Width - (listViewProcess.Columns[1].Width + listViewProcess.Columns[0].Width)) - 32;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}