using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial08 : ITutorial 
    {
        #region ITutorial

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

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial08_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial08_DE_CS"; }

        }

        public string Caption
        {
            get { return "Tutorial08"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Using the Invoker" : "Den Invoker verwenden"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion
    }
}
