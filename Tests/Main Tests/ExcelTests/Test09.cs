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

                Office.COMAddIn addin = (from a in application.COMAddIns where a.ProgId == "NOTestsMain.ExcelTestAddinCSharp" select a).FirstOrDefault();
                if (null == addin || null == addin.Object)
                    return new TestResult(false, DateTime.Now.Subtract(startTime), "NOTestsMain.ExcelTestAddinCSharp or addin.Object not found.", null, "");
                	
                bool addinStatusOkay = false;
                string errorDescription = string.Empty;
                if (null != addin.Object)
                { 
                    COMObject addinProxy = new COMObject(addin.Object);
                    addinStatusOkay = (bool)Invoker.Default.PropertyGet(addinProxy, "StatusOkay");
                    errorDescription = (string)Invoker.Default.PropertyGet(addinProxy, "StatusDescription");
                    addinProxy.Dispose();
                }

                if (addinStatusOkay)
                    return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
                else
                    return new TestResult(false, DateTime.Now.Subtract(startTime), string.Format("NOTestsMain.ExcelTestAddinCSharp Addin Status {0}", errorDescription), null, "");
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
