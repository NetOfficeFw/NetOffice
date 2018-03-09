using System;
using System.Windows.Forms;
using TutorialsBase;
using Excel = NetOffice.ExcelApi;
using NetOffice;

namespace TutorialsCS4
{
    public class Tutorial07 : ITutorial
    {
        public void Run()
        {
            // NetOffice Core supports so-called managed C# dynamic
            // with proxy management services. No need for additional NetOffice Api assemblies.

            // NetOffice want convert a proxy to COMDynamicObject each time if its failed to resolve
            // a corresponding wrapper type.

            // Note: Reference to Microsoft.CSharp is required.

            dynamic application = new COMDynamicObject("Excel.Application");
            application.DisplayAlerts = false;
            var book = application.Workbooks.Add();

            foreach (var sheet in book.Sheets)
            {
                Console.WriteLine(sheet);
            }

            // quit and dispose all open proxies
            application.Quit();
            application.Dispose();

            // -- no proxies open anymore --

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
            get { return Program.DocumentationBase + "Tutorial07_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial07"; }
        }


        public string Description
        {
            get { return "Managed Dynamics"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}