using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.GlobalHelperModules;

namespace TutorialsCS4
{
    public class Tutorial12 : ITutorial
    {
        #region ITutorial

        public void Run()
        {
            // this example demonstrate the global helper module(static class)
            // the module is a vba compatibility workarround and contains static methods and properties from the coresponding Application class.

            // start excel and add a new workbook
            Excel.Application application = new Excel.Application();
            application.Visible = false;
            application.DisplayAlerts = false;
            application.Workbooks.Add();

            // GlobalModule contains the well known globals and is located in NetOffice.ExcelApi.GlobalHelperModules
            // In VB.NET you can do now: ActiveCell.Value = "ActiveCellValue" and this is helpful to bring code from VBA to NetOffice
            GlobalModule.ActiveCell.Value = "ActiveCellValue";

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

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial12_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial12_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial12"; }
        }


        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Globals in NetOffice" : "Globals verwenden"; }
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
