using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperUtils.ProcessKiller
{
    public partial class ProcessKillerControl : UserControl, IUtilsControl
    {
        #region Fields

        private OfficeProcessKiller _processKiller;
        private InfoControl _infoBox;

        private int _currentLanguage;

        #endregion

        #region Construction

        public ProcessKillerControl()
        {
            InitializeComponent();
        }

        public ProcessKillerControl(object anyTag)
        {
            InitializeComponent();
            _processKiller = new OfficeProcessKiller(listViewApps);
            textBoxHotKey.Text = _processKiller.HotKey.ToString();
        }

        #endregion
    
        #region Trigger

        private void listViewApps_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            string appName = e.Item.Text;
            switch (appName)
            {
                case "Excel":
                    _processKiller.Excel = e.Item.Checked;
                    break;
                case "Winword":
                    _processKiller.Word = e.Item.Checked;
                    break;
                case "Outlook":
                    _processKiller.Outlook = e.Item.Checked;
                    break;
                case "PowerPnt":
                    _processKiller.PowerPoint = e.Item.Checked;
                    break;
                case "MsAccess":
                    _processKiller.Access = e.Item.Checked;
                    break;
            }
        }

        private void checkBoxAppsTray_CheckedChanged(object sender, EventArgs e)
        {
            _processKiller.TrayIcon = checkBoxAppsTray.Checked;
        }

        private void checkBoxAppKill_CheckedChanged(object sender, EventArgs e)
        {
            _processKiller.HotKeyEnabled = checkBoxAppKill.Checked;
        }

        private void textBoxHotKey_KeyDown(object sender, KeyEventArgs e)
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
                _processKiller.HotKey = e.KeyCode | Keys.Control | Keys.Alt;
            else if (e.Control)
                _processKiller.HotKey = e.KeyCode | Keys.Control;
            else if (e.Alt)
                _processKiller.HotKey = e.KeyCode | Keys.Alt;
            else
                _processKiller.HotKey = e.KeyCode;
        }

        private void buttonKillApps_Click(object sender, EventArgs e)
        {
            _processKiller.KillProcesses();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (null != _processKiller)
            {
                _processKiller.Dispose();
                _processKiller = null;
            }
        }

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            _infoBox = new InfoControl("ProcessKiller.Info.txt", true);
            this.Controls.Add(_infoBox);
            _infoBox.BringToFront();
            _infoBox.Show();
        }

        #endregion
       
        #region IUtilsControl Members

        public string ControlName
        {
            get 
            {
                return "ProcessKiller";
            }
        }

        public void Activate()
        {
            if (null != _infoBox)
            {
                _infoBox.Hide();
            }
        }

        public void LoadConfiguration(XmlNode configNode)
        {
            if (configNode.ChildNodes.Count == 0)
                configNode.InnerXml = ReadString("ProcessKiller.DefaultConfiguration.xml");

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

            val = configNode.SelectSingleNode("Control/Tray").Attributes[0].Value;
            checkBoxAppsTray.Checked = Convert.ToBoolean(val);

            val = configNode.SelectSingleNode("Control/HotKey").Attributes[1].Value;
            textBoxHotKey.Text = ((Keys)Convert.ToInt32(val)).ToString();
            _processKiller.HotKey= (Keys)Convert.ToInt32(val);

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

            configNode.SelectSingleNode("Control/Tray").Attributes[0].Value = checkBoxAppsTray.Checked.ToString();
            configNode.SelectSingleNode("Control/HotKey").Attributes[0].Value = checkBoxAppKill.Checked.ToString();

            configNode.SelectSingleNode("Control/HotKey").Attributes[1].Value = ((int)_processKiller.HotKey).ToString();
        }

        public void SetLanguage(int id)
        {
            _currentLanguage = id;

            if (0 == id)
            {
  
            }
            else
            {

            }
        }

        public void Release()
        {
            if ((null != _processKiller) && (!this.DesignMode))
                _processKiller.Dispose();
        }

        #endregion

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
    }
}
