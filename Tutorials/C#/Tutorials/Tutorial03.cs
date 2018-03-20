using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public partial class Tutorial03 : UserControl, ITutorial
    {
        #region Fields

        private Excel.Application _application;

        #endregion

        #region Ctor

        public Tutorial03()
        {
            InitializeComponent();
            CreateHandle();
            // update proxy count label
            NetOffice.Core.Default.ProxyCountChanged += delegate (int proxyCount)
            {
                Action<int> update = delegate(int i) { labelProxyCount.Text = i.ToString(); };
                labelProxyCount.Invoke(update, proxyCount);
            };
        }

        #endregion

        #region ITutorial

        public void Run()
        {
            // this tutorial shows you 3 ways in NetOffice to see how many com proxies
            // was currently alive in your application
            //
            // 1.) the property: int NetOffice.Core.ProxyCount
            // 2.) the event: NetOffice.Core.ProxyCountChanged
            // 3.) the events: NetOffice.Core ProxyAdded, ProxyRemoved, ProxyCleared
            //     used from NetOffice.Contribution.Controls.InstanceMonitor

            // Note: Sometimes you may wondering why an instance is disposed.
            // For troubleshooting you can trigger ICOMObject.OnDispose event and see strack trace
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
            instanceMonitor1.Factory = NetOffice.Core.Default;
        }

        public void Disconnect()
        {
            if (null != _application)
            {
                _application.Quit();
                _application.Dispose();
                _application = null;
            }
            instanceMonitor1.Factory = null;
        }

        public string Uri
        {
            get { return Program.DocumentationBase + "Tutorial03_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial03"; }
        }


        public string Description
        {
            get { return "Observable COM proxies"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region Button Trigger

        private void buttonExcel_Click(object sender, EventArgs e)
        {
            if (null == _application)
            {
                // create application
                _application = new Excel.Application();
                _application.DisplayAlerts = false;
                buttonExcel.Text = "Quit Excel";
                buttonWorkbook.Enabled = true;
                buttonAddins.Enabled = true;
                buttonDisposeChildInstances.Enabled = true;
            }
            else
            {
                // dispose application
                _application.Quit();
                _application.Dispose();
                _application = null;
                buttonExcel.Text = "Start Excel";
                buttonWorkbook.Enabled = false;
                buttonAddins.Enabled = false;
                buttonDisposeChildInstances.Enabled = false;
            }
        }

        private void buttonWorkbook_Click(object sender, EventArgs e)
        {
            // 2 new proxies, the workbooks proxy(implicit) and the new workbook from Add()
            if (null != _application)
                _application.Workbooks.Add();
        }

        private void buttonAddins_Click(object sender, EventArgs e)
        {
            if (null != _application)
            {
                // 1 new enumerator proxy and 1 new proxy for any Addin
                foreach (Excel.AddIn item in _application.AddIns)
                    Console.WriteLine(item.Name);
            }
        }

        private void buttonDisposeChildInstances_Click(object sender, EventArgs e)
        {
            // dispose all child instances from application
            _application.DisposeChildInstances();
        }

        #endregion
    }
}