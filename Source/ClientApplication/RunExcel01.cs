using System;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using Contribution = NetOffice.OfficeApi.Tools.Contribution;

namespace ClientApplication
{
    internal class RunExcel01
    {
        internal void Run()
        {
            Excel.Application app = null;
            try
            {
                Settings.Default.PerformanceTrace.Alert += new PerformanceTrace.PerformanceAlertEventHandler(PerformanceTrace_Alert);
                Settings.Default.PerformanceTrace["ExcelApi"].Enabled = true;
                Settings.Default.PerformanceTrace["ExcelApi"].IntervalMS = 0;

                app = new Excel.Application();
                app.Visible = true;
                Contribution.CommonUtils utils = new Contribution.CommonUtils(app, typeof(Form1).Assembly);
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