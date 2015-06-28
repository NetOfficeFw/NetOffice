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

                // Criteria 1
                // Enable performance trace in excel generaly. set interval limit to 100 to see all actions there need >= 100 milliseconds
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi"].Enabled = true;
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi"].IntervalMS = 100;

                // Criteria 2
                // Enable additional performance trace for all members of Range in excel. set interval limit to 20 to see all actions there need >=20 milliseconds
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi", "Range"].Enabled = true;
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi", "Range"].IntervalMS = 20;

                // Criteria 3
                // Enable additional performance trace for WorkSheet Range property in excel. set interval limit to 0 to see all calls anywhere
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi", "Worksheet", "Range"].Enabled = true;
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi", "Worksheet", "Range"].IntervalMS = 0;

                // Criteria 4
                // Enable additional performance trace for Range this[] indexer in excel. set interval limit to 0 to see all calls anywhere
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi", "Range", "_Default"].Enabled = true;
                NetOffice.Settings.Default.PerformanceTrace["ExcelApi", "Range", "_Default"].IntervalMS = 0;

                Excel.Application application = new Excel.Application();
                application.DisplayAlerts = false;
                Excel.Workbook book = application.Workbooks.Add();
                Excel.Worksheet sheet = book.Sheets.Add() as Excel.Worksheet;
                for (int i = 1; i <= 5; i++)
                {
                    Excel.Range range = sheet.Range("A" + i.ToString());
                    range.Value = "Test123";
                    range[1, 1].Value = "Test234";
                }

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
