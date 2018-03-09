using System;
using System.Linq;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using NetOffice.Extensions;
using NetOffice.Extensions.Invoker;
using System.Collections.Generic;
using NetOffice.CollectionsGeneric;

namespace TutorialsCS4
{
    public class Tutorial11 : ITutorial
    {
        public void Run()
        {
            // Best practice to write own IEnumerable<T> extensions.
            // See sample extension at the end of these file.

            // NetOffice spend some extensions on IEnumerable<T> you may know from Linq2Objects.
            // These extensions take care to free unused/unwanted COM Proxies immediately.
            // However, the extensions doesnt works like Linq which means calling the result
            // execute the method on demand. Its just ordinary extensions.

            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;
            application.Workbooks.Add();

            // Here we use "First()" and "FirstOrDefault()" and the Invoker extension "Property" because Sheets is an untyped collection
            Excel.Worksheet sheet = application.Workbooks.First().Sheets.FirstOrDefault(e => e.Property<string>("Name") == "Sheet1") as Excel.Worksheet;
            if (null != sheet)
            {
                sheet.Cells[1, 1].Value = "Test123";
                sheet.Cells[5, 5].Value = "Test123";
                sheet.Cells[10, 10].Value = "Test123";

                // iterate over 10x10 used range and return the 3 cells we filled
                // Linq2Objects would create 101 new proxies(10x10 + enumerator) here without any chance to free them.
                // In NetOffice exensions - you have just 4 new managed proxies.
                var ranges = sheet.UsedRange.Where(e => e.Value != null);

                // doing the same here again with the tutorial sample extension (scroll down)
                ranges = sheet.UsedRange.AllCellsWithValues();
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
            get { return Program.DocumentationBase + "Tutorial11_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial11"; }
        }

        public string Description
        {
            get { return "Extensions and IEnumerable<T>"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }

    }

    internal static class Tutorial11Sample
    {
        // -- Best Practice sample extension to create extensions for IEnumerable<T> in NetOffice
        //
        // In order to prevent ambiguous conflicts
        // you need to target NetOffice.CollectionsGeneric.IEnumerableProvider<T>
        // All collections in NetOffice implement these interface
        public static IEnumerable<Excel.Range> AllCellsWithValues(this IEnumerableProvider<Excel.Range> source)
        {
            List<Excel.Range> result = new List<NetOffice.ExcelApi.Range>();
            ICOMObject enumerator = source.GetComObjectEnumerator(null);
            try
            {
                foreach (Excel.Range item in source.FetchVariantComObjectEnumerator(source as ICOMObject, enumerator))
                {
                    if (item.Value != null)
                        result.Add(item);
                    else
                        item.Dispose();
                }
                return result;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
            }
        }
    }
}