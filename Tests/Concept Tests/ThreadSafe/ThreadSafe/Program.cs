using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Text;
using Excel = NetOffice.ExcelApi;

namespace ThreadSafe
{
    public class Program
    {
        private static Excel.Application _application;
         
        public static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("--- EnableThreadSafe Test --- ");

                for (int i = 0; i < 100; i++)
                {
                    Console.WriteLine("Do Test {0}", i+1);
                    DoTest();
                }

                Console.WriteLine("Test passed.");
                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Test failed.");
                Console.WriteLine(exception.Message);
                Console.ReadKey();
            }
        }

        private static void DoTest()
        {
            _application = new Excel.Application();
            Excel.Workbook book = _application.Workbooks.Add();

            WaitHandle[] waitHandles = new WaitHandle[3];

            Thread thread1 = new Thread(new ParameterizedThreadStart(Thread1Method));
            Thread thread2 = new Thread(new ParameterizedThreadStart(Thread2Method));
            Thread thread3 = new Thread(new ParameterizedThreadStart(Thread3Method));

            ManualResetEvent mre1 = new ManualResetEvent(false);
            ManualResetEvent mre2 = new ManualResetEvent(false);
            ManualResetEvent mre3 = new ManualResetEvent(false);

            waitHandles[0] = mre1;
            waitHandles[1] = mre2;
            waitHandles[2] = mre3;

            thread1.Start(mre1);
            thread2.Start(mre2);
            thread3.Start(mre3);

            WaitHandle.WaitAll(waitHandles);

            _application.Quit();
            _application.Dispose();
        }

        // here comes a DRY issue but its just a test okay ;)

        private static void Thread1Method(object mre)
        {
            Excel.Worksheet sheet = _application.ActiveSheet as Excel.Worksheet;
            foreach (Excel.Range range in sheet.Range("A1:B200"))
            {
                Excel.Workbook book = range.Application.ActiveWorkbook;
                foreach (object item in book.Sheets)
                {
                    Excel.Worksheet otherSheet = item as Excel.Worksheet;
                    Excel.Range rng = otherSheet.Cells[1, 1];
                }
                range.Dispose();   
            }
            (mre as ManualResetEvent).Set();
        }

        private static void Thread2Method(object mre)
        {
            Excel.Worksheet sheet = _application.ActiveSheet as Excel.Worksheet;
            foreach (Excel.Range range in sheet.Range("A1:B200"))
            {
                Excel.Workbook book = range.Application.ActiveWorkbook;
                foreach (object item in book.Sheets)
                {
                    Excel.Worksheet otherSheet = item as Excel.Worksheet;
                    Excel.Range rng = otherSheet.Cells[1, 1];
                }
            }
            (mre as ManualResetEvent).Set();
        }

        private static void Thread3Method(object mre)
        {
            Excel.Worksheet sheet = _application.ActiveSheet as Excel.Worksheet;
            foreach (Excel.Range range in sheet.Range("A1:B200"))
            {
                Excel.Workbook book = range.Application.ActiveWorkbook;
                foreach (object item in book.Sheets)
                {
                    Excel.Worksheet otherSheet = item as Excel.Worksheet;
                    Excel.Range rng = otherSheet.Cells[1, 1];
                }
            }
            sheet.DisposeChildInstances();
            (mre as ManualResetEvent).Set();
        }
    }
}
