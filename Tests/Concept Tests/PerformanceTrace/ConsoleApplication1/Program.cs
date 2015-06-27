using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = NetOffice.ExcelApi;
using Outlook = NetOffice.OutlookApi;
using PowerPoint = NetOffice.PowerPointApi;
using Word = NetOffice.WordApi;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 
                Console.WriteLine("NetOffice Utils Concept Test");
                Console.WriteLine("0 Milliseconds trace values is not a bug - its just to fast\r\n");

                NetOffice.Settings.Default.PerformanceTrace.Alert += new NetOffice.PerformanceTrace.PerformanceAlertEventHandler(PerformanceTrace_Alert);

                // Test 1:
                // Enable performance trace in excel generaly. set interval limit to 0 to see all actions
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi"].Enabled = true;
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi"].IntervalMS = 0;

                // Test 2:
                // Enable performance trace for range in excel. set interval limit to 100 to see all actions there need 100 or more milliseconds
                // NetOffice.Settings.Default.PerformanceTrace["Excel", "Range"].Enabled = true;
                // NetOffice.Settings.Default.PerformanceTrace["Excel", "Range"].IntervalMS = 100;

                // Test 3: 
                // Enable performance trace for range this[] indexer in excel. set interval limit to 10 to see all actions there need 10 or more milliseconds
                // NetOffice.Settings.Default.PerformanceTrace["Excel", "Range", "Item"].Enabled = true;
                // NetOffice.Settings.Default.PerformanceTrace["Excel", "Range", "Item"].IntervalMS = 10;

                Excel.Application application = new Excel.Application();
                application.DisplayAlerts = false;
                Excel.Workbook book = application.Workbooks.Add();
                Excel.Worksheet sheet = book.Sheets.Add() as Excel.Worksheet;
                for (int i = 1; i <= 5; i++)
                    sheet.Range("A" + i.ToString()).Value = "Test123";

                application.Quit();
                application.Dispose();

                Console.WriteLine("\r\nTest passed");
                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                Console.ReadKey();
            }
        }

        private static void PerformanceTrace_Alert(NetOffice.PerformanceTrace sender, NetOffice.PerformanceTrace.PerformanceAlertEventArgs args)
        {
            Console.WriteLine("{0} {1}:{2} in {3} Milliseconds ({4} Ticks)", args.CallType, args.EntityName, args.MethodName, args.TimeElapsedMS, args.Ticks);
        }
    }
}
