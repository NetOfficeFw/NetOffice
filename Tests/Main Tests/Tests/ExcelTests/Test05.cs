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
    /// charts and datasource
    /// </summary>
    public class Test05 : ITestPackage
    {
        #region ITestPackage Member

        public bool DoTest(string logFilePath)
        {
            Core.DebugConsole.FileName = System.IO.Path.Combine(logFilePath, "ExcelTests.Test05.log");
            Core.DebugConsole.AppendTimeInfoEnabled = true;
            Core.DebugConsole.Mode = LateBindingApi.Core.ConsoleMode.LogFile;

            Excel.Application application = null;
            try
            {
                LateBindingApi.Core.Factory.Initialize();

                // start excel and turn off msg boxes
                application = new Excel.Application();
                application.DisplayAlerts = false;

                // add a new workbook
                Excel.Workbook workBook = application.Workbooks.Add();
                Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

                // we need some data to display
                Excel.Range dataRange = PutSampleData(workSheet);

                // create a nice diagram
                Excel.ChartObject chart = ((Excel.ChartObjects)workSheet.ChartObjects()).Add(70, 100, 375, 225);
                chart.Chart.SetSourceData(dataRange);

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

        private static Excel.Range PutSampleData(Excel.Worksheet workSheet)
        {
            workSheet.Cells[2, 2].Value = "Datum";
            workSheet.Cells[3, 2].Value = DateTime.Now.ToShortDateString();
            workSheet.Cells[4, 2].Value = DateTime.Now.ToShortDateString();
            workSheet.Cells[5, 2].Value = DateTime.Now.ToShortDateString();
            workSheet.Cells[6, 2].Value = DateTime.Now.ToShortDateString();

            workSheet.Cells[2, 3].Value = "Columns1";
            workSheet.Cells[3, 3].Value = 25;
            workSheet.Cells[4, 3].Value = 33;
            workSheet.Cells[5, 3].Value = 30;
            workSheet.Cells[6, 3].Value = 22;

            workSheet.Cells[2, 4].Value = "Column2";
            workSheet.Cells[3, 4].Value = 25;
            workSheet.Cells[4, 4].Value = 33;
            workSheet.Cells[5, 4].Value = 30;
            workSheet.Cells[6, 4].Value = 22;

            workSheet.Cells[2, 5].Value = "Column3";
            workSheet.Cells[3, 5].Value = 25;
            workSheet.Cells[4, 5].Value = 33;
            workSheet.Cells[5, 5].Value = 30;
            workSheet.Cells[6, 5].Value = 22;

            return workSheet.get_Range("$B2:$E6");
        }
    }
}
