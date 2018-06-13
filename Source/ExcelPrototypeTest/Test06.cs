using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    internal class Test06
    {
        internal void Run()
        {
            dynamic application = new COMDynamicObject("Excel.Application");
            application.DisplayAlerts = false;
            var book = application.Workbooks.Add();

            Excel.Workbook convertedBook = book as Excel.Workbook;
            Console.WriteLine("book converted {0}", null != convertedBook);

            foreach (var sheet in book.Sheets)
            {
                Console.WriteLine(sheet);
            }

            application.Quit();
            application.Dispose();
        }
    }
}
