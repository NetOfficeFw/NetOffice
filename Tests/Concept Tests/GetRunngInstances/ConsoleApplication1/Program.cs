using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = NetOffice.ExcelApi;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("NetOffice Concept Test - Excel.Application.GetActiveInstances(){0}", Environment.NewLine);
            Excel.Application[] apps = Excel.Application.GetActiveInstances();

            Console.WriteLine("{0} active Excel instance(s) found.{1}", apps.Length, Environment.NewLine);

            foreach (Excel.Application app in Excel.Application.GetActiveInstances())
            {
                string workbooks = string.Empty;
                foreach (Excel.Workbook openBook in app.Workbooks)
                    workbooks += openBook.Name + " ";

                Console.WriteLine("Excel.Application {0} has open workbooks:{1}", app.Hwnd, workbooks);
                app.Dispose();
            }

            Console.WriteLine("{0}Press any key...", Environment.NewLine);
            Console.ReadKey();
        }

    }
}
