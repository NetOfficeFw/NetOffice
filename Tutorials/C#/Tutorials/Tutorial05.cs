using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;

namespace TutorialsCS4
{
    public class Tutorial05 : ITutorial
    {
        public void Run()
        {
            // this is a simple demonstration how to convert unkown types at runtime

            // start application
            Excel.Application application = new Excel.Application();
            application.Visible = false;
            application.DisplayAlerts = false;

            foreach (Office.COMAddIn item in application.COMAddIns)
            {
                // the application property is an unkown COM object
                // we need a simple cast at runtime
                Excel.Application hostApp = item.Application as Excel.Application;

                // do some sample stuff
                string hostAppName = hostApp.Name;
                bool hostAppVisible = hostApp.Visible;
            }

            // quit and dispose excel
            application.Quit();
            application.Dispose();

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
            get { return Program.DocumentationBase + "Tutorial05_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial05"; }
        }

        public string Description
        {
            get { return "Understanding unknown types"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}