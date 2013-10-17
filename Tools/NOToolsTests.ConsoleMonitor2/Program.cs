using System;
using System.Collections.Generic;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace NOToolsTests.ConsoleMonitor2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("NOToolsTests.ConsoleMonitor2{0}Press any key to start.", Environment.NewLine);
            Console.ReadKey();
            Console.WriteLine("Running...");

            NetOffice.DebugConsole.Default.EnableSharedOutput = true;
            Excel.Application application = new Excel.Application();
            application.Visible = false;
            application.DisplayAlerts = false;
            application.Workbooks.Add();
            application.Quit();
            application.Dispose();

            Console.WriteLine("Press any key...");
            Console.ReadKey();
        }
    }
}
