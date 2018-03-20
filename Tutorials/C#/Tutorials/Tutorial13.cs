using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial13 : ITutorial
    {
        public void Run()
        {
            // this examples shows a special method to ask at runtime for a particular method oder property
            // morevover you can enable the option NetOffice.Settings.EnableSafeMode. 
            // NetOffice checks(cache supported) for any method or property you call and
            // throws a EntitiyNotSupportedException if missing

            // create new instance
            Excel.Application application = new Excel.Application();

            // any reference type in NetOffice implements the EntityIsAvailable method.
            // you check here your property or method is available.

            // we check the support for 2 properties  at runtime
            bool enableLivePreviewSupport = application.EntityIsAvailable("EnableLivePreview");
            bool openDatabaseSupport = application.Workbooks.EntityIsAvailable("OpenDatabase");

            string result = "Excel Runtime Check: " + Environment.NewLine;
            result += "Support EnableLivePreview: " + enableLivePreviewSupport.ToString() + Environment.NewLine;
            result += "Support OpenDatabase:      " + openDatabaseSupport.ToString() + Environment.NewLine;

            // quit and dispose
            application.Quit();
            application.Dispose();

            HostApplication.ShowMessage(result);
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
            get { return Program.DocumentationBase + "Tutorial13_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial13"; }
        }

        public string Description
        {
            get { return  "Version-independent development"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
