using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public partial class Tutorial02 : ITutorial
    {
        IHost _hostApplication;
         
        #region ITutorial Member

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
            get { return _hostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial02_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial02_DE_CS"; }

        }

        public string Caption
        {
            get { return "Tutorial02"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Using Dispose & DisposeChildInstances" : "Verwenden von Dispose und DisposeChildInstances"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
