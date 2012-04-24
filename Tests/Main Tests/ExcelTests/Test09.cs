using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using System.Linq;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace ExcelTestsCSharp
{
    public class Test09 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test09"; }
        }

        public string Description
        {
            get { return "Check for loaded Addin"; }
        }

        public string OfficeProduct
        {
            get { return "Excel"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Excel.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                application = new NetOffice.ExcelApi.Application();
                application.Visible = true;
                application.DisplayAlerts = false;
                application.Workbooks.Add();
                Excel.Worksheet sheet = application.Workbooks[1].Sheets[1] as Excel.Worksheet;

                Office.COMAddIn addin = (from a in application.COMAddIns where a.ProgId == "ExcelAddinCSharp.TestAddin" select a).FirstOrDefault();
                if(null == addin)
                    return new TestResult(false, DateTime.Now.Subtract(startTime), "COMAddin ExcelAddinCSharp.TestAddin not found.", null, "");
      
                COMObject addinProxy = new COMObject(null, addin.Object);
                bool ribbonIsOkay = (bool)Invoker.PropertyGet(addinProxy, "RibbonUIPassed");
                bool taskPaneIsOkay = (bool)Invoker.PropertyGet(addinProxy, "TaskPanePassed");
                addinProxy.Dispose();

                if( ribbonIsOkay && taskPaneIsOkay)
                    return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
                else
                    return new TestResult(false, DateTime.Now.Subtract(startTime), string.Format("Ribbon:{0} TaskPane:{1}", ribbonIsOkay, taskPaneIsOkay), null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    application.Quit();
                    application.Dispose();
                }
            }
        }

        #endregion
    }
}
