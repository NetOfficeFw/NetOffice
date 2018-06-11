using System;
using System.Windows.Forms;
using TutorialsBase;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial09 : ITutorial
    {
        public void Run()
        {
            // NetOffice instances implements the IClonable interface
            // and deal with underlying proxies as well

            Excel.Application application = new Excel.ApplicationClass();
            application.DisplayAlerts = false;
            Excel.Workbook book = application.Workbooks.Add();

            // clone the book
            Excel.Workbook cloneBook = book.Clone() as Excel.Workbook;

            // dispose the origin book keep the underlying proxy alive
            // until the clone is disposed
            book.Dispose();

            // alive and works even the origin book is disposed
            foreach (Excel.Worksheet sheet in cloneBook.Sheets)
            {
                Console.WriteLine(sheet);
            }

            application.Quit();
            application.Dispose();

            HostApplication.ShowFinishDialog();
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public void Disconnect()
        {

        }

        public string Uri
        {
            get { return Program.DocumentationBase + "Tutorial09_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial09"; }
        }


        public string Description
        {
            get { return "Cloning Instances"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
