using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using System.Globalization;
using ExampleBase;

using NetOffice;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace ExcelExamplesCS4
{
    class Example01 : IExample
    {
        IHost _hostApplication;
        
        #region IExample Member

        public void RunExample()
        {
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

            // draw back color and perform the BorderAround method
            workSheet.Range("$B2:$B5").Interior.Color = ToDouble(Color.DarkGreen);
            workSheet.Range("$B2:$B5").BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic);

            // draw back color and border the range explicitly
            workSheet.Range("$D2:$D5").Interior.Color = ToDouble(Color.DarkGreen);
            workSheet.Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlDouble;
            workSheet.Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].Weight = 4;
            workSheet.Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].Color = ToDouble(Color.Black);

            workSheet.Cells[1, 1].Value = "We have 2 simple shapes created.";

            // save the book 
            string fileExtension = GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example01{1}", _hostApplication.RootDirectory, fileExtension);
            workBook.SaveAs(workbookFile);

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
            get { return _hostApplication.LCID == 1033 ? "Example01" : "Beispiel01"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Background Colors and Borders for Cells" : "Hintergrundfarben und Rahmen in Zellen"; }
        }

        public UserControl Panel 
        {
            get { return null; }
        }

        #endregion

        #region Helper

        /// <summary>
        /// Translate a color to double
        /// </summary>
        /// <param name="color">expression to convert</param>
        /// <returns>color</returns>
        private static double ToDouble(System.Drawing.Color color)
        {
            uint returnValue = color.B;
            returnValue = returnValue << 8;
            returnValue += color.G;
            returnValue = returnValue << 8;
            returnValue += color.R;
            return returnValue;
        }

        /// <summary>
        /// returns the valid file extension for the instance. for example ".xls" or ".xlsx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Excel.Application application)
        {
            double Version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (Version >= 12.00)
                return ".xlsx";
            else
                return ".xls";
        }

        #endregion
    }
}
