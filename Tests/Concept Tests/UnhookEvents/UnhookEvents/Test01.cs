using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace UnhookEvents
{
    public class Test01
    {
        private int _activatedCounter = 0;

        public void Proceed()
        {
            using (var application = new Excel.ApplicationClass())
            {
                application.Settings.ExceptionMessageBehavior = ExceptionMessageHandling.DiagnosticsAndInnerMessage;
                application.Settings.EnableAutomaticQuit = true;
                application.Visible = true;
                application.DisplayAlerts = false;

                application.SheetActivateEvent += Application_SheetActivateEvent;

                Console.WriteLine("Press any key to unhook SheetActivateEvent");
                Console.ReadKey();

                application.SheetActivateEvent -= Application_SheetActivateEvent;

                Console.WriteLine("Press any key to close.");
                Console.ReadKey();
            }
        }

        private void Application_SheetActivateEvent(ICOMObject sh)
        {
            _activatedCounter++;
            Console.WriteLine("Sheet activated. {0} times.", _activatedCounter);
        }
    }
}
