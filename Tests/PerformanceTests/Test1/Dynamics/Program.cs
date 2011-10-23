using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.IO;

namespace Dynamics
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("C# Dynamics Performance Test - 5000 Cells");
            Console.WriteLine("Write simple text.");

            // start excel, and get a new sheet reference
            dynamic excelApplication = CreateExcelApplication();
            dynamic books = excelApplication.Workbooks;
            dynamic book = books.Add(Missing.Value);
            dynamic sheets = book.Worksheets;
            dynamic sheet = sheets.Add();

            // do test 10 times
            List<MarshalByRefObject> comReferencesList = new List<MarshalByRefObject>();
            List<TimeSpan> timeElapsedList = new List<TimeSpan>();
            for (int i = 1; i <= 10; i++)
            {
                DateTime timeStart = DateTime.Now;
                for (int y = 1; y <= 5000; y++)
                {
                    string rangeAdress = "$A" + y.ToString();
                    dynamic range = sheet.Range[rangeAdress];
                    range.Value = "value";
                    comReferencesList.Add(range as MarshalByRefObject);
                }
                TimeSpan timeElapsed = DateTime.Now - timeStart;

                // display info and dispose references
                Console.WriteLine("Time Elapsed: {0}", timeElapsed);
                timeElapsedList.Add(timeElapsed);
                foreach (var item in comReferencesList)
                    Marshal.ReleaseComObject(item);
                comReferencesList.Clear();
            }

            // display info & log to file
            TimeSpan timeAverage = AppendResultToLogFile(timeElapsedList, "Test1-Dynamics.log");
            Console.WriteLine("Time Average: {0}{1}Press any key...", timeAverage, Environment.NewLine);
            Console.Read();

            // release & quit
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(books);

            excelApplication.Quit();
            Marshal.ReleaseComObject(excelApplication);

        }

        /// <summary>
        /// creates a new excel application
        /// </summary>
        /// <returns></returns>
        static dynamic CreateExcelApplication()
        {
            // start excel
            Type excelType = System.Type.GetTypeFromProgID("Excel.Application");
            dynamic excelApplication = System.Activator.CreateInstance(excelType);
            excelApplication.DisplayAlerts = false;
            excelApplication.Interactive = false;
            excelApplication.ScreenUpdating = false;
            return excelApplication;
        }

        /// <summary>
        /// writes list items to a logile and append average of items at the end
        /// </summary>
        /// <param name="timeElapsedList">a list with log results</param>
        /// <param name="fileName">name of logfile in current assembly folder</param>
        /// <returns>average of timeElapsedList</returns>
        static TimeSpan AppendResultToLogFile(List<TimeSpan> timeElapsedList, string fileName)
        {
            TimeSpan timeSummary = TimeSpan.Zero;
            string logFile = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), fileName);

            if (File.Exists(logFile))
                File.Delete(logFile);

            foreach (TimeSpan item in timeElapsedList)
            {
                timeSummary += item;
                string logFileAppend = item.ToString() + Environment.NewLine;
                File.AppendAllText(logFile, logFileAppend, Encoding.UTF8);
            }

            TimeSpan timeAverage = TimeSpan.FromTicks(timeSummary.Ticks / timeElapsedList.Count);
            File.AppendAllText(logFile, "Time Average: " + timeAverage.ToString(), Encoding.UTF8);
            return timeAverage;
        }
    }
}
