using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using Core = NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace ExcelTestsCSharp
{
    /// <summary>
    /// foreach enumeration
    /// </summary>
    public class Test07 : ITestPackage
    {
        #region TestPackage Member
         
        public string Name
        {
            get { return "Test07"; }
        }

        public string Description
        {
            get { return "Range ForEach enumeration"; }
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
                application.DisplayAlerts = false;
                application.Workbooks.Add();
                Excel.Worksheet sheet = application.Workbooks[1].Sheets[1] as Excel.Worksheet;

                foreach (Excel.Range item in sheet.Range("$A1:$B100"))
                    item.Value = DateTime.Now;

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
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
