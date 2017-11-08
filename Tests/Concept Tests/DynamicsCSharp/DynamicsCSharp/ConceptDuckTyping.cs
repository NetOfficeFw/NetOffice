using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using IExcel = NetOffice.IExcelApi;

namespace DynamicsCSharp
{
    internal class ConceptDuckTyping
    {
        internal void Test()
        {
            IExcel.IApplication application = Core.Default.CreateDuckObject<IExcel.IApplication>();
            application.Visible = true;
            application.DisplayAlerts = false;
            application.NewWorkbookEvent += Application_NewWorkbookEvent;

            application.Workbooks.Add().Sheets.Add(null, application.Workbooks[1].Sheets[3]);
            
            foreach (IExcel.IWorkbook book in application.Workbooks)
            {
                Console.WriteLine(book.Name);
                foreach (IExcel.IWorksheet sheet in book.Sheets)
                {
                    Console.WriteLine(sheet.Name);
                    sheet.Cells[3, 3].Value = "Test 123";
                }
            }
            
            application.Workbooks[1].SaveAs(Path.Combine(Environment.CurrentDirectory, "File.xlsx"));

            application.Quit();
            application.Dispose();

            Console.WriteLine("Press any key...");
            Console.ReadKey();
        }

        private void Application_NewWorkbookEvent(IExcel.IWorkbook wb)
        {
            Console.WriteLine("NewWorkbookEvent {0}", wb.Name);
        }
    }
}