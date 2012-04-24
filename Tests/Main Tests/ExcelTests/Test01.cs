using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using Core = NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace ExcelTestsCSharp
{
    public class Test01 : ITestPackage
    {
        #region TestPackage Member
         
        public string Name
        {
            get { return "Test01"; }
        }

        public string Description
        {
            get { return "Simple Range.Value call for a Worksheet."; }
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

                for (int i = 1; i <= 200; i++)
                {
                    sheet.Cells[i, 1].Value = string.Format("Test {0}", i);
                    sheet.Range(string.Format("$B{0}", i)).Value = 42.3;
                }

                return new TestResult(true, DateTime.Now.Subtract(startTime), "",null, "");
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
