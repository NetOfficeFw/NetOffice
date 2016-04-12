using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = NetOffice.ExcelApi;
using Outlook = NetOffice.OutlookApi;
using PowerPoint = NetOffice.PowerPointApi;
using Word = NetOffice.WordApi;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("NetOffice Utils Concept Test");

                TestWord();
                TestExcel();
                TestOutlook();
                TestPowerPoint();

                Console.WriteLine("Test passed");
                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                Console.ReadKey();
            }
        }

        private static void TestWord()
        {
            Console.WriteLine("Test Word Application Utils");
            
            Word.Application application = new Word.Application();
            application.DisplayAlerts = Word.Enums.WdAlertLevel.wdAlertsNone;
            application.Documents.Add();

            Word.Tools.Utils.CommonUtils utils = new Word.Tools.Utils.CommonUtils(application);
            int hwnd = utils.Application.TryGetMainWindowHandle(application.Documents[1]);

            application.Quit();
            application.Dispose();

            if (0 == hwnd)
                throw new Exception("Cant resolve word hwnd");
        }

        private static void TestPowerPoint()
        {
            Console.WriteLine("Test PowerPoint Application Utils");

            PowerPoint.Application application = new PowerPoint.Application();
            application.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
            application.Presentations.Add();

            PowerPoint.Tools.Utils.CommonUtils utils = new PowerPoint.Tools.Utils.CommonUtils(application);
            int hwnd = utils.Application.HWND;

            application.Quit();
            application.Dispose();

            if (0 == hwnd)
                throw new Exception("Cant resolve powerpoint hwnd");
        }

        private static void TestOutlook()
        {
            Console.WriteLine("Test Outlook Application Utils");

            Outlook.Application application = new Outlook.Application();
            Outlook.Tools.Utils.CommonUtils utils = new Outlook.Tools.Utils.CommonUtils(application);

            bool visible1 = utils.Application.Visible;
            application.Session.GetDefaultFolder(Outlook.Enums.OlDefaultFolders.olFolderInbox).Display();
            System.Threading.Thread.Sleep(3000);
            bool visible2 = utils.Application.Visible;

            application.Quit();
            application.Dispose();

            if(!(false == visible1 && true == visible2))
                throw new Exception("Unexpected outlook visibility");
        }

        private static void TestExcel()
        {
            Console.WriteLine("Test Excel File Utils");

            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;
            Excel.Tools.Utils.CommonUtils utils = new Excel.Tools.Utils.CommonUtils(application);

            string fileName = utils.File.Combine("C:\\MyFiles", "Test01", Excel.Tools.DocumentFormat.Normal);

            application.Quit();
            application.Dispose();

            if (utils.ApplicationIs2007OrHigher)
            {
                if ("C:\\MyFiles\\Test01.xlsx" != fileName)
                    throw new Exception("Unexpected excel filename");
            }
            else
            {
                if ("C:\\MyFiles\\Test01.xls" != fileName)
                    throw new Exception("Unexpected excel filename");
            }            
        }
    }
}
