using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = NetOffice.ExcelApi;

namespace OperatorOverloadsCSharp
{
    class Program
    {
        //
        // Test: NetOffice.Settings.EnableOperatorOverlads (true by default)
        //
        static void Main(string[] args)
        {
            // start excel
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            // add 3 workbooks
            Excel.Workbook book1 = application.Workbooks.Add();
            Excel.Workbook book2 = application.Workbooks.Add();
            Excel.Workbook book3 = application.Workbooks.Add();

            // check "==" operator
            if (application.ActiveWorkbook == book1)
                Console.WriteLine("Book 1 is ActiveWorkbook");
            else if (application.ActiveWorkbook == book2)
                Console.WriteLine("Book 2 is ActiveWorkbook");
            else if (application.ActiveWorkbook == book3)
                Console.WriteLine("Book 3 is ActiveWorkbook");
            else
                Console.WriteLine("Operator Overload failed");

            // check "!=" operator
            if (application.ActiveWorkbook != book1)
                Console.WriteLine("Book 1 is not ActiveWorkbook");
            if (application.ActiveWorkbook != book2)
                Console.WriteLine("Book 2 is not ActiveWorkbook");
            if (application.ActiveWorkbook != book3)
                Console.WriteLine("Book 3 is not ActiveWorkbook");

            // close and dispose
            application.Quit();
            application.Dispose();

            Console.WriteLine("Press any key...");
            Console.Read();
        }
    }
}
