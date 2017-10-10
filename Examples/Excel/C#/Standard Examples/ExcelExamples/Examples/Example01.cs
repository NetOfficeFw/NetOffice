using System;
using System.Drawing;
using System.Windows.Forms;
using ExampleBase;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.ExcelApi.Tools.Contribution;

namespace ExcelExamplesCS4
{
    /// <summary>
    /// Example 1 - Background Colors and Borders for Cells
    /// </summary>
    internal class Example01 : IExample
    {
        public void RunExample()
        {
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // create a utils instance, no need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(excelApplication);

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

            // draw back color and call the BorderAround method
            workSheet.Range("$B2:$B5").Interior.Color = utils.Color.ToDouble(Color.DarkGreen);
            workSheet.Range("$B2:$B5").BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic);

            // draw back color and border the range explicitly
            workSheet.Range("$D2:$D5").Interior.Color = utils.Color.ToDouble(Color.DarkGreen);
            workSheet.Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlDouble;
            workSheet.Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].Weight = 4;
            workSheet.Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].Color = utils.Color.ToDouble(Color.Black);

            workSheet.Cells[1, 1].Value = "We have 2 simple shapes created.";

            // save the book - utils want build the filename for us
            string workbookFile = utils.File.Combine(HostApplication.RootDirectory, "Example01", DocumentFormat.Normal);
            workBook.SaveAs(workbookFile);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            // show end dialog
            HostApplication.ShowFinishDialog(null, workbookFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return "Example01"; }
        }

        public string Description
        {
            get { return "Background Colors and Borders for Cells"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
