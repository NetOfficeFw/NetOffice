using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using NetOffice.Running;

namespace TutorialsCS4
{
    public class Tutorial14 : ITutorial
    {
        public void Run()
        {
            // In some situations you want use NetOffice with an already running application.
            // this tutorial shows how its possible.

            // 1)
            //
            // GetActiveInstance take the first instance in memory
            Excel.Application application = ProxyService.GetActiveInstance<Excel.Application>();
            if (null != application)
                application.Dispose();

            // 2)
            //
            // GetActiveInstances takes all instances in memory
            var applications = ProxyService.GetActiveInstances<Excel.Application>();
            applications.Dispose();

            // 3)
            //
            // Use special ctor to try access a running application first
            // and if its failed create a new application
            application = new Excel.ApplicationClass(new Core(), true);
            // quit only if its a new application
            if (!application.FromProxyService)
                application.Quit();
            application.Dispose();

            // 4)
            //
            // Creates instance from interop proxy
            Type interopType = Type.GetTypeFromProgID("Excel.Application");
            object proxy = Activator.CreateInstance(interopType);
            application = COMObject.Create<Excel.Application>(proxy);
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
            get { return Program.DocumentationBase + "Tutorial14_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial14"; }
        }

        public string Description
        {
            get { return "Accessing running applications"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
