using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = NetOffice.ExcelApi;
using PowerPoint = NetOffice.PowerPointApi;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            GetExcelActiveInstances();
            GetActivePowerPointInstance();
            Console.WriteLine("{0}Press any key...", Environment.NewLine);
            Console.ReadKey();
        }

        static void GetExcelActiveInstances()
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
        }

        static void GetActivePowerPointInstance()
        {
            PowerPoint.Application application = null;
            try
            {
                NetOffice.Settings.Default.UseExceptionMessage = NetOffice.ExceptionMessageHandling.CopyInnerExceptionMessageToTopLevelException;
                Console.WriteLine("NetOffice Concept Test - PowerPoint.Application.GetActiveInstance(){0}", Environment.NewLine);
                application = PowerPoint.Application.GetActiveInstance(false);
                if (null != application)
                {
                    Console.WriteLine("Current PowerPoint Application Visibility: {0}", application.Visible);
                }
                else 
                {
                    Console.WriteLine("No PowerPoint Application running.");
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("An error has occured. {0}", exception.Message);
            }
            finally
            {
                if (null != application)
                {
                    application.Dispose();
                    application = null;
                }
            }
        }
    }
}
