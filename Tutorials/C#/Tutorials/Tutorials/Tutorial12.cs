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
    public partial class Tutorial12 : ITutorial
    {
        IHost _hostApplication;

        #region ITutorial Member

        public void Run()
        {
            // start excel and add a new workbook
            Excel.Application application = new Excel.Application();
            application.Visible = false;
            application.DisplayAlerts = false;
            application.Workbooks.Add();

            // GlobalModule contains the well known globals and is located in NetOffice.ExcelApi.GlobalHelperModules
            GlobalModule.ActiveCell.Value = "ActiveCellValue";

            // quit and dispose excel
            application.Quit();
            application.Dispose();

            _hostApplication.ShowFinishDialog();
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public void Disconnect()
        {

        }

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return _hostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial12_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial12_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial12"; }
        }


        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Globals in NetOffice" : "Globals verwenden"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
