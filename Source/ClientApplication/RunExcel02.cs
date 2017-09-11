using System;
using NetOffice;

namespace ClientApplication
{
    internal class RunExcel02
    {
        internal void Run()
        {
            try
            {
                NetOffice.Settings.Default.EnableAutomaticQuit = true;
                using (dynamic application = new COMDynamicObject("Excel.Application"))
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
