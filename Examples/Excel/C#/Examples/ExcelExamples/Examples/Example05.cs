using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using System.Globalization;
using ExampleBase;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.VBIDEApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace ExcelExamplesCS4
{
    class Example05 : IExample
    {
        IHost _hostApplication;

        #region IExample Member

        public void RunExample()
        {      
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

            // we need some data to display
            Excel.Range dataRange = PutSampleData(workSheet);

            // create a nice diagram
            Excel.ChartObject chart = ((Excel.ChartObjects)workSheet.ChartObjects()).Add(70, 100, 375, 225);
            chart.Chart.SetSourceData(dataRange);

            // save the book 
            string fileExtension = GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example05{1}", _hostApplication.RootDirectory, fileExtension);
            workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, NetOffice.ExcelApi.Enums.XlSaveAsAccessMode.xlExclusive);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            // show dialog for the user(you!)
            _hostApplication.ShowFinishDialog(null, workbookFile);
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example05" : "Beispiel05"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Working with Charts" : "Arbeiten mit Charts"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Helper

        private static Excel.Range PutSampleData(Excel.Worksheet workSheet)
        {
            workSheet.Cells[2, 2].Value = "Datum";
            workSheet.Cells[3, 2].Value = DateTime.Now.ToShortDateString();
            workSheet.Cells[4, 2].Value = DateTime.Now.ToShortDateString();
            workSheet.Cells[5, 2].Value = DateTime.Now.ToShortDateString();
            workSheet.Cells[6, 2].Value = DateTime.Now.ToShortDateString();

            workSheet.Cells[2, 3].Value = "Columns1";
            workSheet.Cells[3, 3].Value = 25;
            workSheet.Cells[4, 3].Value = 33;
            workSheet.Cells[5, 3].Value = 30;
            workSheet.Cells[6, 3].Value = 22;

            workSheet.Cells[2, 4].Value = "Column2";
            workSheet.Cells[3, 4].Value = 25;
            workSheet.Cells[4, 4].Value = 33;
            workSheet.Cells[5, 4].Value = 30;
            workSheet.Cells[6, 4].Value = 22;

            workSheet.Cells[2, 5].Value = "Column3";
            workSheet.Cells[3, 5].Value = 25;
            workSheet.Cells[4, 5].Value = 33;
            workSheet.Cells[5, 5].Value = 30;
            workSheet.Cells[6, 5].Value = 22;

            return workSheet.Range("$B2:$E6");
        }

        /// <summary>
        /// returns the valid file extension for the instance. for example ".xls" or ".xlsx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Excel.Application application)
        {
            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".xlsx";
            else
                return ".xls";
        }

        #endregion
    }
}
