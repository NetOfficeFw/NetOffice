using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using NetOffice.Exceptions;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    internal class Test07
    {
        internal void Run()
        {
            MyCore core = new MyCore();
            core.Settings.EnableAutomaticQuit = true;
            core.Settings.ForceApplicationVersionProviders = true;
            using (Excel.Application application = new Excel.ApplicationClass(core))
            {
                application.DisplayAlerts = false;
                var workbooks = application.Workbooks;
                var book = workbooks.Add();
                var sheet = book.Sheets.Add() as Excel.Worksheet;

                try
                {
                    sheet.Range("NONSENS").Value = "value";
                }
                catch (NetOfficeCOMException exception)
                {
                    Console.WriteLine("NetOfficeCOMException, NetOffice Version:{0} Application Version:{1}",
                        exception.NetOfficeVersion, exception.ApplicationVersion);
                }
                catch (Exception exception)
                {
                    Console.WriteLine("Unexpected Exception {0}", exception);
                }               
            }
        }
    }
}
