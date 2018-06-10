using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    internal class Test02
    {
        internal void Run()
        {
            MyCore core = new MyCore();
            core.Settings.EnableAutomaticQuit = true;
            using (Excel.Application application = new Excel.ApplicationClass(core))
            {
                application.DisplayAlerts = false;
                var workbooks = application.Workbooks;
                var book = workbooks.Add();
                book.Sheets.Add();
    
                string instanceFriendlyName = application.InstanceFriendlyName;
                string instanceComponentName = application.InstanceComponentName;

                Console.WriteLine("InstanceFriendlyName:{0}", instanceFriendlyName);
                Console.WriteLine("InstanceComponentName:{0}", instanceComponentName);
            }
        }
    }
}
