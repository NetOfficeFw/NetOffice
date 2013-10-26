using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = NetOffice.ExcelApi;

namespace OnDisposeTest
{
    class Program
    {
        private static bool CancelDispose;

        static void Main(string[] args)
        {
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;
            application.Visible = false;
            application.OnDispose += new NetOffice.OnDisposeEventHandler(application_OnDispose);
            application.Workbooks.Add();

            CancelDispose = true;

            application.Dispose(); // cancel the first dispose

            CancelDispose = false;

            application.Quit();
            application.Dispose();
        }

        static void application_OnDispose(NetOffice.OnDisposeEventArgs eventArgs)
        {
            eventArgs.Cancel = CancelDispose;
        }
    }
}
