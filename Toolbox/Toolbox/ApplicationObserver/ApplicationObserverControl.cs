using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ApplicationObserver
{
    public partial class ApplicationObserverControl : UserControl, IToolboxControl
    {
        #region Fields

        private OfficeApplicationObserver _applicationObserver;
        private int                       _currentLanguageID = 1033;
     
        #endregion

        #region Construction

        public ApplicationObserverControl()
        {
            try
            {
                InitializeComponent();
                if (!DesignMode)
                {
                    _applicationObserver = new OfficeApplicationObserver(listViewApps);
                    textBoxHotKey.Text = _applicationObserver.HotKey.ToString();
                    _applicationObserver.InstanceRunningCountChanged += new EventHandler(_applicationObserver_InstanceRunningCountChanged);
                    _applicationObserver.AllProcessesChanged += new EventHandler(_applicationObserver_AllProcessesChanged);
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
               errorForm.ShowDialog(this);
            }
            
        }

        #endregion
        
        #region Properties
       
        private new bool DesignMode
        {
            get
            {
                return (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "devenv");
            }
        }
       
        #endregion

        #region Trigger
         
        void _applicationObserver_AllProcessesChanged(object sender, EventArgs e)
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
            }
            catch (Exception exception)
            {

                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        void _applicationObserver_InstanceRunningCountChanged(object sender, EventArgs e)
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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

                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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

                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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

                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            try
            {
                InfoControl infoBox = new InfoControl("ApplicationObserver.Info" + _currentLanguageID.ToString() + ".rtf", true);
                this.Controls.Add(infoBox);
                infoBox.BringToFront();
                infoBox.Show();
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void checkBoxShowQuestion_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                _applicationObserver.ShowQuesionBeforeKill = checkBoxShowQuestion.Checked;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion
       
        #region IUtilsControl Members

        public new void KeyDown(KeyEventArgs e)
        { 
        }

        public string ControlName
        {
            get 
            {
                return "ApplicationObserver";
            }
        }

        public string ControlCaption
        {
            get { return "Application Observer"; }
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
            get
            {
                return ReadImageFromRessource("Icon.png");
            }
        }

        public void Activate()
        {
            
        }

        public void LoadComplete()
        {
           
        }

        public void LoadConfiguration(XmlNode configNode)
        {
            if (configNode.ChildNodes.Count == 0)
                configNode.InnerXml = ReadString("ApplicationObserver.IconsAndConfig.DefaultConfiguration.xml");

            string val = "";

            val = configNode.SelectSingleNode("Control/Excel").Attributes[0].Value;
            listViewApps.Items[0].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Control/Winword").Attributes[0].Value;
            listViewApps.Items[1].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Control/Outlook").Attributes[0].Value;
            listViewApps.Items[2].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Control/PowerPnt").Attributes[0].Value;
            listViewApps.Items[3].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Control/MsAccess").Attributes[0].Value;
            listViewApps.Items[4].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Control/Project").Attributes[0].Value;
            listViewApps.Items[5].Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Control/Visio").Attributes[0].Value;
            listViewApps.Items[6].Checked = Convert.ToBoolean(val);
             
            val = configNode.SelectSingleNode("Control/Tray").Attributes[0].Value;
            checkBoxAppsTray.Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Control/HotKey").Attributes[1].Value;
            textBoxHotKey.Text = ((Keys)Convert.ToInt32(val)).ToString();
            _applicationObserver.HotKey = (Keys)Convert.ToInt32(val);

            val = configNode.SelectSingleNode("Control/HotKey").Attributes[0].Value;
            checkBoxAppKill.Checked = Convert.ToBoolean(val);
        }

        public void SaveConfiguration(XmlNode configNode)
        {
            configNode.SelectSingleNode("Control/Excel").Attributes[0].Value = listViewApps.Items[0].Checked.ToString();
            configNode.SelectSingleNode("Control/Winword").Attributes[0].Value = listViewApps.Items[1].Checked.ToString();
            configNode.SelectSingleNode("Control/Outlook").Attributes[0].Value = listViewApps.Items[2].Checked.ToString();
            configNode.SelectSingleNode("Control/PowerPnt").Attributes[0].Value = listViewApps.Items[3].Checked.ToString();
            configNode.SelectSingleNode("Control/MsAccess").Attributes[0].Value = listViewApps.Items[4].Checked.ToString();
            configNode.SelectSingleNode("Control/Project").Attributes[0].Value = listViewApps.Items[5].Checked.ToString();
            configNode.SelectSingleNode("Control/Visio").Attributes[0].Value = listViewApps.Items[6].Checked.ToString();

            configNode.SelectSingleNode("Control/Tray").Attributes[0].Value = checkBoxAppsTray.Checked.ToString();
            configNode.SelectSingleNode("Control/HotKey").Attributes[0].Value = checkBoxAppKill.Checked.ToString();

            configNode.SelectSingleNode("Control/HotKey").Attributes[1].Value = ((int)_applicationObserver.HotKey).ToString();
        }

        public void SetLanguage(int id)
        {
            _currentLanguageID = id;
            _applicationObserver.CurrentLanguageID = id;
            Translator.TranslateControls(this, "ApplicationObserver.MessageTable.txt", _currentLanguageID);
        }

        public new void Dispose()
        {
            if ((null != _applicationObserver) && (!this.DesignMode))
            {
                _applicationObserver.Dispose();
                _applicationObserver = null;
            }
            base.Dispose();
        }

        #endregion
        
        #region Helper

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

        private static string ReadString(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            System.IO.StreamReader textStreamReader = null;
            try
            {
                string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                ressourcePath = assemblyName + "." + ressourcePath;
                ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
                if (ressourceStream == null)
                    throw (new System.IO.IOException("Error accessing resource Stream."));

                textStreamReader = new System.IO.StreamReader(ressourceStream);
                if (textStreamReader == null)
                    throw (new System.IO.IOException("Error accessing resource File."));

                string text = textStreamReader.ReadToEnd();
                return text;
            }
            catch (Exception exception)
            {
                throw (exception);
            }
            finally
            {
                if (null != textStreamReader)
                    textStreamReader.Close();
                if (null != ressourceStream)
                    ressourceStream.Close();
            }
        }

        private static Image ReadImageFromRessource(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            ressourcePath = assemblyName + ".ApplicationObserver." + ressourcePath;
            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            Bitmap newIcon = new Bitmap(ressourceStream);
            return newIcon;
        }

        #endregion
    }
}
