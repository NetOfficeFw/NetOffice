using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using NetOffice;
using NetOffice.Attributes;

namespace ClientApplication
{
    internal class RunExcel02
    {
        internal void Run()
        {
            try
            {
                Type excelType = System.Type.GetTypeFromProgID("Excel.Application", true);
                object interopProxy = Activator.CreateInstance(excelType);

                NetOffice.Settings.Default.EnableAutomaticQuit = true;
                using (dynamic application = new COMDynamicObject(interopProxy))
                {
                    application.Visible = true;
                    application.Workbooks.Add();
                  
                    dynamic dynamicBooks = new COMDynamicObject(application.Workbooks.UnderlyingObject);
                    var book = dynamicBooks[1];

                    bool bookActive = application.ActiveWorkbook == application.Workbooks[1];
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }
     }
}
