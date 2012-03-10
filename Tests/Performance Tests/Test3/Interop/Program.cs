using System;
using System.IO;
using System.Collections.Generic; 
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Interop
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Interop Assemblies Performance Test - 15000 Cells.");
            Console.WriteLine("Write simple text, change Font, NumberFormat, WrapText and add a comment.");

            // start excel, and get a new sheet reference
            Excel.Application excelApplication = CreateExcelApplication();
            Excel.Workbooks books = excelApplication.Workbooks;
            Excel.Workbook book = books.Add(Missing.Value);
            Excel.Sheets sheets = book.Worksheets;
            Excel.Worksheet sheet = sheets.Add() as Excel.Worksheet;

            // do test 10 times
            List<MarshalByRefObject> comReferencesList = new List<MarshalByRefObject>();
            List<TimeSpan> timeElapsedList = new List<TimeSpan>();
            for (int i = 1; i <= 10; i++)
            {
                sheet.UsedRange.ClearComments();
                DateTime timeStart = DateTime.Now;
                for (int y = 1; y <=15000; y++)
                {
                    string rangeAdress = "$A" + y.ToString();
                    Excel.Range cellRange = sheet.get_Range(rangeAdress);
                    cellRange.Value = "value";
                    Excel.Font font = cellRange.Font;
                    font.Name = "Verdana";
                    cellRange.NumberFormat = "@";
                    cellRange.WrapText = true;
                    Excel.Comment sampleComment = cellRange.AddComment("Sample Comment");
                    comReferencesList.Add(font as MarshalByRefObject);
                    comReferencesList.Add(sampleComment as MarshalByRefObject);
                    comReferencesList.Add(cellRange as MarshalByRefObject);
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
            TimeSpan timeAverage = AppendResultToLogFile(timeElapsedList, "Test3-Interop.log");
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
