using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    internal class Test01
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
            }

            var typeCache = core.Cache.GetTypeCache();
            Console.WriteLine("--Start Type Cache Log--");
            foreach (var item in typeCache)
            {
                Console.WriteLine("Cache Item Factory:{0}\r\nComponent:{1}\r\nType:{2}\r\nContract:{3}\r\nImplementation:{4}\r\n", 
                    item.Factory.FactoryName, item.ComponentId, item.TypeId, item.Contract.FullName, item.Implementation.FullName);
            }
            Console.WriteLine("--End Type Cache Log--");
        }
    }
}
