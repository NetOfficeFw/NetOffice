using System;
using System.Windows.Forms;
using TutorialsBase;
using Excel = NetOffice.ExcelApi;
using static NetOffice.ExcelApi.ModulesLegacy.ApplicationModule;

namespace TutorialsCS4
{
    public class Tutorial16 : ITutorial
    {
        public void Run()
        {
            // this example demonstrate the global helper module(static class)
            // the module is a vba compatibility workarround and contains static methods and properties from the coresponding Application class.

            // start excel and add a new workbook
            Excel.Application application = new Excel.ApplicationClass();
            application.Visible = false;
            application.DisplayAlerts = false;
            application.Workbooks.Add();

            // ApplicationModule contains the well known globals and its located in NetOffice.$XXXApi.ModulesLegacy
            // This helpful to bring code from VBA to NetOffice
            ActiveCell.Value = "ActiveCellValue";

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
            get { return Program.DocumentationBase + "Tutorial16_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial16"; }
        }

        public string Description
        {
            get { return "Globals in NetOffice"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
