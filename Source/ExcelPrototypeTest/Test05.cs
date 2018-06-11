using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    internal class Test05
    {
        internal void Run()
        {
            var applications = NetOffice.ProxyService.GetActiveInstances<Excel.Application>();
            Console.WriteLine("Found {0} applications", applications.Count);
            applications.Dispose();

            var application = NetOffice.ProxyService.GetActiveInstance<Excel.Application>();
            Console.WriteLine("Found application {0}", null != application);
            application?.Dispose();
        }
    }
}
