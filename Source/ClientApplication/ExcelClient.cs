using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;

using Utils = NetOffice.OfficeApi.Tools.Utils;


namespace ClientApplication
{
    internal class ExcelClient
    {
        internal void Test()
        {
            Excel.Application app = null;
            try
            {
                Settings.Default.PerformanceTrace.Alert += new PerformanceTrace.PerformanceAlertEventHandler(PerformanceTrace_Alert);
                Settings.Default.PerformanceTrace["ExcelApi"].Enabled = true;
                Settings.Default.PerformanceTrace["ExcelApi"].IntervalMS = 0;

                app = new Excel.Application();
                Utils.CommonUtils utils = new Utils.CommonUtils(app, typeof(Form1).Assembly);
                app.DisplayAlerts = false;
                Excel.Workbook book = app.Workbooks.Add();
                Excel.Worksheet sheet = book.Sheets[1] as Excel.Worksheet;
                sheet.Cells[1, 1].Value = "This is a sample value";
                sheet.Protect();

                utils.Dialog.SuppressOnAutomation = false;
                utils.Dialog.SuppressOnHide = false;
                utils.Dialog.ShowDiagnostics(true);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
            }
            finally
            {
                if (null != app)
                {
                    app.Quit();
                    app.Dispose();
                }
            }
        }

        private void PerformanceTrace_Alert(PerformanceTrace sender, PerformanceTrace.PerformanceAlertEventArgs e)
        {
            Console.WriteLine("Call {4} => {0}:{1} passed in {2} milliseconds ({3} Ticks)", e.EntityName, e.MethodName, e.TimeElapsedMS, e.Ticks, e.CallType);
        }
    }
}
