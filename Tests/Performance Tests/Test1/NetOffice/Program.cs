using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = NetOffice.ExcelApi;

namespace NetOffice
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("NetOffice Release 1.3 Performance Test - 5000 Cells.");
            Console.WriteLine("Write simple text.");

            // start excel, and get a new sheet reference
            Excel.Application excelApplication = CreateExcelApplication();
            Excel.Worksheet sheet = excelApplication.Workbooks.Add().Worksheets.Add() as Excel.Worksheet;
          
            // do test 10 times
            List<TimeSpan> timeElapsedList = new List<TimeSpan>();
            for (int i = 1; i <= 10; i++)
            {
                DateTime timeStart = DateTime.Now;
                for (int y = 1; y <= 5000; y++)
                {
                    string rangeAdress = "$A" + y.ToString();
                    sheet.get_Range(rangeAdress).Value = "value";
                }
                TimeSpan timeElapsed = DateTime.Now - timeStart;

                // display info and dispose references
                Console.WriteLine("Time Elapsed: {0}", timeElapsed);
                timeElapsedList.Add(timeElapsed);
                sheet.DisposeChildInstances();
            }
          
            // display info & log to file
            TimeSpan timeAverage = AppendResultToLogFile(timeElapsedList, "Test1-NetOffice.log");
            Console.WriteLine("Time Average: {0}{1}Press any key...", timeAverage, Environment.NewLine);
            Console.Read();

            // release & quit
            excelApplication.Quit();
            excelApplication.Dispose();
        }

        /// <summary>
        /// creates a new excel application
        /// </summary>
        /// <returns></returns>
        static Excel.Application CreateExcelApplication()
        {
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            excelApplication.Interactive = false;
            excelApplication.ScreenUpdating = false;
            return excelApplication;
        }
        
        /// <summary>
        /// writes list items to a logile and append average of items at the end
        /// </summary>
        /// <param name="timeElapsedList">a list with log results</param>
        /// <param name="fileName">name of logfile in current assembly folder </param>
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
