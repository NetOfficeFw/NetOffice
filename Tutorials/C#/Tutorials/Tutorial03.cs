using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public partial class Tutorial03 : UserControl, ITutorial
    {
        #region Fields
        
        IHost _hostApplication;
        Excel.Application _application;

        #endregion

        #region Ctor
        
        public Tutorial03()
        {
            InitializeComponent();

            // add event trigger to ProxyCountChanged event
            NetOffice.Core.Default.ProxyCountChanged += new Core.ProxyCountChangedHandler(ProxyCountChanged);
        }

        #endregion

        #region ITutorial

        public void Run()
        { 
            // this example shows you both ways in NetOffice to see how many com proxies
            // was currently alive in your application
            //
            // 1.) the static property: int NetOffice.Factory.ProxyCount
            // 2.) the static event: NetOffice.Factory.ProxyCountChanged
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public void Disconnect()
        {
            if (null != _application)
            {
                _application.Quit();
                _application.Dispose();
            }
        }

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return _hostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial03_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial03_DE_CS"; }

        }

        public string Caption
        {
            get { return "Tutorial03"; }
        }


        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Observable COM Proxy Count" : "Die Anzahl COM Proxies überwachen"; }
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
                // start application
                _application = new Excel.Application();
                _application.DisplayAlerts = false;
                buttonExcel.Text = "Quit Excel";
                buttonWorkbook.Enabled = true;
                buttonAddins.Enabled = true;
                buttonAddRemoveWorkbook.Enabled = true;
            }
            else
            {
                _application.Quit();
                _application.Dispose();
                _application = null;
                buttonExcel.Text = "Start Excel";
                buttonWorkbook.Enabled = false;
                buttonAddins.Enabled = false;
                buttonAddRemoveWorkbook.Enabled = false;
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

        private void buttonAddRemoveWorkbook_Click(object sender, EventArgs e)
        {
            // add a new worbook and a new worksheet to the workbook
            // the worksheet is a child proxy from worbook, after dispose the workbook
            // creates 4 new proxies
            // the open proxy count is the same as before

            int proxyCount = NetOffice.Core.Default.ProxyCount;

            Excel.Workbook book = _application.Workbooks.Add();
            book.Worksheets.Add();

            int proxyCountAfterCreate = NetOffice.Core.Default.ProxyCount;

            // dispose all child instances from application
            _application.DisposeChildInstances();

            int proxyCountAfterDispose = NetOffice.Core.Default.ProxyCount;

            string message = string.Format(
                                           "ProxyCount before create is {0}\r\n" +
                                           "ProxyCount after create is {1}\r\n" +
                                           "ProxyCount after dispose all childs from application is {2}", proxyCount, proxyCountAfterCreate, proxyCountAfterDispose);

            MessageBox.Show(message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region ProxyCountChanged Trigger

        // its possible the event comes from a different thread, the method is an invoke helper to avoid a CrossThreadException
        private void UpdateLabel()
        {
            labelProxyCount.Text = labelProxyCount.Tag as string;
        }

        void ProxyCountChanged(int proxyCount)
        {
            if (labelProxyCount.InvokeRequired)
            {
                labelProxyCount.Tag = proxyCount.ToString();
                labelProxyCount.Invoke(new MethodInvoker(UpdateLabel));
            }
            else
            {

            }
                labelProxyCount.Text = proxyCount.ToString();
        }

        #endregion
    }
}
