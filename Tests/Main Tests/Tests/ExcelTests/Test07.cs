using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using Core = LateBindingApi.Core;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace ExcelTests
{
    /// <summary>
    /// foreach enumeration
    /// </summary>
    public class Test07 : ITestPackage
    {
        #region TestPackage Member

        public bool DoTest(string logFilePath)
        {
            Core.DebugConsole.FileName = System.IO.Path.Combine(logFilePath, "ExcelTests.Test07.log");
            Core.DebugConsole.AppendTimeInfoEnabled = true;
            Core.DebugConsole.Mode = LateBindingApi.Core.ConsoleMode.LogFile;

            Excel.Application application = null;
            try
            {
                LateBindingApi.Core.Factory.Initialize();

                application = new NetOffice.ExcelApi.Application();
                application.DisplayAlerts = false;
                application.Workbooks.Add();
                Excel.Worksheet sheet = application.Workbooks[1].Sheets[1] as Excel.Worksheet;

                foreach (Excel.Range item in sheet.Range("$A1:$B100"))
                    item.Value = DateTime.Now;                                     

                return true;
            }
            catch (Exception exception)
            {
                string message = exception.Message;
                Console.WriteLine("An error occured{1}{1}{0}", message, Environment.NewLine);
                return false;
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
