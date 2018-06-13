using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using NetOffice.Exceptions;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    internal class Test08
    {
        internal void Run()
        {
            MyCore core = new MyCore();
            core.Settings.EnableAutomaticQuit = true;
            using (Excel.Application application = new Excel.ApplicationClass(core))
            {
                application.WorkbookActivateEvent += Application_WorkbookActivateEvent;
                application.DisplayAlerts = false;
                var book = application.Workbooks.Add();
                var sheet = book.Sheets.Add() as Excel.Worksheet;               
            }
        }

        private void Application_WorkbookActivateEvent(Excel.Workbook wb)
        {
            Console.WriteLine("WorkbookActivateEvent has been called.");
        }
    }
}
