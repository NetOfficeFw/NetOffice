using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperUtils.RegistryWatcher
{
    public partial class RegistryAlarmControl : UserControl, IUtilsControl 
    {
        public RegistryAlarmControl()
        {
            InitializeComponent();
        }

        public RegistryAlarmControl(object anyTag)
        {
            InitializeComponent();
        }

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            InfoControl infoBox = new InfoControl("RegistryAlarm.Info.txt", true);
            this.Controls.Add(infoBox);
            infoBox.BringToFront();
            infoBox.Show();
        }

        #region IUtilsControl Members

        public string ControlName
        {
            get { return "RegistryAlarm"; }
        }

        public void Activate()
        {

        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
          
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
           
        }

        public void Release()
        {
           
        }

        #endregion

        private void toolStripAdd_Click(object sender, EventArgs e)
        {

        }

        private void toolStripRemove_Click(object sender, EventArgs e)
        {

        }
    }
}
