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
    /// simple Range.Value call on worksheet
    /// </summary>
    public class Test01 : ITestPackage
    {
        #region TestPackage Member

        public bool DoTest(string logFilePath)
        {
            Core.DebugConsole.FileName = System.IO.Path.Combine(logFilePath, "ExcelTests.Test01.log");
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

                for (int i = 1; i <= 200; i++)
                {
                    sheet.Cells[i, 1].Value = string.Format("Test {0}", i);
                    sheet.Range(string.Format("$B{0}", i)).Value = 42.3;
                }

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
