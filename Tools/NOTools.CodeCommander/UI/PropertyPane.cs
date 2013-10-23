using System;
using System.Collections.Generic;
using System.Windows.Forms;
using NetOffice;
using NetOffice.OfficeApi.Tools;
using NOTools.CodeCommander.Logic;

namespace NOTools.CodeCommander.UI
{
    public partial class PropertyPane : UserControl
    {
        public PropertyPane()
        {
            InitializeComponent();
        }

        internal OfficeApplicationManager ApplicationHandler { get; private set; }

        public void OnConnection(COMObject application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            ApplicationHandler = new OfficeApplicationManager(application);
            AvailableProxy[] proxies = ApplicationHandler.GetAvailableProxies();
            if (proxies.Length > 0)
            {
                comboBoxTarget.DataSource = proxies;
                comboBoxTarget.SelectedIndex = 0;
            }
        }

        public void OnDisconnection()
        {

        }

        private void comboBoxTarget_SelectedValueChanged(object sender, EventArgs e)
        {
            if (null != propertyGridHostProperties.SelectedObject)
            {
                COMObject oldSelectedObject = propertyGridHostProperties.SelectedObject as COMObject;
                propertyGridHostProperties.SelectedObject = null;
                propertyGridHostProperties.Refresh();
                oldSelectedObject.DisposeChildInstances();
            }

            propertyGridHostProperties.SelectedObject = ApplicationHandler.GetSelectedProxy(comboBoxTarget.SelectedIndex);
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            propertyGridHostProperties.Refresh();
        }
    }
}
