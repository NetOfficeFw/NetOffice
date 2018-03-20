using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial15 : ITutorial
    {
        public void Run()
        {
            // this example demonstrate the NetOffice low-level interface for latebinding calls

            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;
            application.Workbooks.Add();

            Excel.Worksheet sheet = (Excel.Worksheet)application.Workbooks[1].Worksheets[1];
            Excel.Range sampleRange = sheet.Cells[1, 1];

            // we set the COMVariant ColorIndex from Font of ouer sample range with the invoker class
            Invoker.Default.PropertySet(sampleRange.Font, "ColorIndex", 1);

            // creates a native unmanaged ComProxy with the invoker an release immediately
            object comProxy = Invoker.Default.PropertyGet(application, "Workbooks");
            Marshal.ReleaseComObject(comProxy);

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
            get { return Program.DocumentationBase + "Tutorial15_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial15"; }
        }

        public string Description
        {
            get { return "Using the Invoker"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
