using System;
using System.Collections.Generic;
using System.Text;

using Excel = NetOffice.ExcelApi;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            // init api
            LateBindingApi.Core.Factory.Initialize();

            Console.WriteLine("Write 1 million cells in excel.");

            Excel.Application application = new NetOffice.ExcelApi.Application();
            application.DisplayAlerts = false;
            application.Interactive = false;
            application.ScreenUpdating = false;

            application.Workbooks.Add();

            Excel.Worksheet workSheet = (Excel.Worksheet)application.Workbooks[1].Worksheets[1];

            // row
            for (int i = 1; i <= 10000; i++)
            {
                // column
                for (int y = 1; y <= 100; y++)
                {
                    Excel.Range cells = workSheet.Cells;
                    Excel.Range range = cells[i, y];
                    range.Value = "TestValue";
                    range.Dispose();
                    cells.Dispose();
                }

                if (i % 100 == 0)
                    Console.WriteLine("{0} Cells written.", (i * 100));
            }

            // quit and dispose
            application.Quit();
            application.Dispose();

            Console.WriteLine("Done!");
        }
    }
}
