using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial02 : ITutorial
    {
        public void Run()
        {
            // this example shows you another dispose method: DisposeChildInstances
            // this means all child proxies from an instance

            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            Excel.Workbook book = application.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets.Add();

            /*
             * we have 5 created proxies now in proxy table as follows
             *
             * Application
             *   + Workbooks
             *     + Workbook
             *        + Worksheets
             *            + Worksheet
            */


            // we dispose the child instances from book
            book.DisposeChildInstances();

            /*
            * we have 3 created proxies now, the childs from book are disposed
            *
            * Application
            *   + Workbooks
            *     + Workbook
            */

            application.Quit();
            application.Dispose();

            // the Dispose() call for application release the instance and created childs Workbooks and Workbook

            HostApplication.ShowFinishDialog();
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public void Disconnect()
        {

        }

        public string Uri
        {
            get { return Program.DocumentationBase + "Tutorial02_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial02"; }
        }

        public string Description
        {
            get { return "Using Dispose & DisposeChildInstances"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}