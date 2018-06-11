using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using NetOffice.Extensions.Conversion;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    internal class Test04
    {
        internal void Run()
        {
            MyCore core = new MyCore();
            core.ObjectActivator.RegisterType(typeof(Excel.Application), typeof(MyExcelApplication));
            core.ObjectActivator.RegisterType(typeof(Excel.Range), typeof(MyExcelRange));
            core.Settings.EnableAutomaticQuit = true;
            using (Excel.Application application = COMObject.Create<Excel.Application>(core))
            {
                application.DisplayAlerts = false;
                var workbooks = application.Workbooks;
                var book = workbooks.Add();
                var sheet = book.Sheets.Add().To<Excel.Worksheet>();

                bool visible = application.Visible;
                application.Visible = true;

                try
                {
                    application.Visible = false;
                }
                catch (ArgumentException)
                {
                    Console.WriteLine("MyExcelApplication prevent us to make excel invisible.");
                }

                var range = sheet.Range("$A1");
                range.Value = null;

                var value = range.Value;
                Console.WriteLine("We set null for range but MyExcelRange change it to {0}.", value);
            }
        }
    }
}
