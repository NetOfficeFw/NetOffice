using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Xml;
using System.Text;
using System.Windows.Forms;

using NetOffice.DeveloperUtils.ProcessKiller;
using NetOffice.DeveloperUtils.RegistryBrowser;

namespace NetOffice.DeveloperUtils
{
    public partial class Form1 : Form
    {
        #region Fields
        
        NotifyIcon  _notify;
        string[]    _args;
        XmlDocument _configFile;
        List<IUtilsControl> _controls = new List<IUtilsControl>();

        #endregion

        #region Construction

        public Form1()
        {
            InitializeComponent();
        } 

        public Form1(string[] args)
        {
            InitializeComponent();

            ProcessKillerControl newControl = new ProcessKillerControl(null);
            tabControlMain.TabPages.Add(newControl.ControlName);
            tabControlMain.TabPages[tabControlMain.TabPages.Count - 1].Controls.Add(newControl);
            tabControlMain.TabPages[tabControlMain.TabPages.Count - 1].Tag = newControl;
            newControl.Dock = DockStyle.Fill;
            _controls.Add(newControl);

            RegistryBrowserControl newControl2 = new RegistryBrowserControl(null);
            tabControlMain.TabPages.Add(newControl2.ControlName);
            tabControlMain.TabPages[tabControlMain.TabPages.Count - 1].Controls.Add(newControl2);
            tabControlMain.TabPages[tabControlMain.TabPages.Count - 1].Tag = newControl2;
            newControl2.Dock = DockStyle.Fill;
            _controls.Add(newControl2);

            LoadConfiguration();
            SetupTrayIcon(true);
            _args = args;
        }

        #endregion

        #region Config Methods

        private void LoadConfiguration()
        {
            string filePath = Path.Combine(Application.StartupPath, "DeveloperUtils.Settings.xml");            
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
            
            foreach (IUtilsControl item in _controls)
                item.LoadConfiguration(_configFile.SelectSingleNode("Application/Controls/" + item.ControlName));
            
        }

        private void SaveConfiguration()
        {
            XmlNode appNode = _configFile.SelectSingleNode("Application");
            appNode.Attributes.Item(0).Value = checkBoxStartAppMinimized.Checked.ToString();
            appNode.Attributes.Item(1).Value = checkBoxMinimizeToTray.Checked.ToString();
      
            foreach (IUtilsControl item in _controls)
                item.SaveConfiguration(_configFile.SelectSingleNode("Application/Controls/" + item.ControlName));

            string filePath = Path.Combine(Application.StartupPath, "DeveloperUtils.Settings.xml");
            _configFile.Save(filePath);
        }

        #endregion

        #region Methods

        private void SetupTrayIcon(bool init)
        {
            if (true == init)
            {
                _notify = new NotifyIcon();
                _notify.Icon = this.Icon;
                _notify.Text = "DeveloperUtils";
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfiguration();
            foreach (IUtilsControl item in _controls)
                item.Release();
 
            SetupTrayIcon(false);
        }
        
        void _notify_Click(object sender, EventArgs e)
        {
            this.Show();
            _notify.Visible = false;
            this.WindowState = FormWindowState.Normal; 
            this.Activate();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if ((FormWindowState.Minimized == this.WindowState) && (true == checkBoxMinimizeToTray.Checked))
            {
                _notify.Visible = true;
                this.Hide();
            }
            else
            {
                _notify.Visible = false;
            }
        }
      
        private void Form1_Shown(object sender, EventArgs e)
        {
            bool minimize = false;

            if (true == checkBoxStartAppMinimized.Checked)
                minimize = true;

            foreach (string item in _args)
                if (true == item.Trim().ToLower().Equals("-min", StringComparison.InvariantCultureIgnoreCase))
                    minimize = true;

            if(minimize)
                this.WindowState = FormWindowState.Minimized;
        }
        
        private void richTextBoxDescription_LinkClicked(object sender, LinkClickedEventArgs e)
        {

        }

        #endregion
  
        private void tabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            TabPage currentTab = tabControlMain.TabPages[tabControlMain.SelectedIndex];
            IUtilsControl control = currentTab.Tag as IUtilsControl;
            if (null != control)
            {
                control.Activate();
            }
        }

        private void linkLabelHomepage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(linkLabelHomepage.Text);
        }
    }
}
