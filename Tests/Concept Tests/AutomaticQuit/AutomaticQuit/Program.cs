using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.Extensions;

namespace AutomaticQuit
{
    class Program
    {
        static void Main(string[] args)
        {
            Core core = new Core();
            core.Settings.ForceApplicationVersionProviders = false;
            core.Settings.EnableAutomaticQuit = true;

            using (var application = new Excel.ApplicationClass(core))
            {
                application.Visible = true;
                var sheet = application.
                                Workbooks.
                                Add().
                                Worksheets[1].
                                SetProperty<Excel.Worksheet>("Name", "MyWorksheet");

                application.Quit();
                core.Invoker.Method(application, "Foo123");
            }
        }
    }
}