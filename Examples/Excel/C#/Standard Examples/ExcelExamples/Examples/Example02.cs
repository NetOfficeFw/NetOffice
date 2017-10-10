using System.Drawing;
using System.Windows.Forms;
using ExampleBase;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.ExcelApi.Tools.Contribution;

namespace ExcelExamplesCS4
{
    /// <summary>
    /// Example 2 - Font Attributes and Alignment for Cells
    /// </summary>
    internal class Example02 : IExample
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

            // font action
            workSheet.Range("A1").Value = "Arial Size:8 Bold Italic Underline";
            workSheet.Range("A1").Font.Name = "Arial";
            workSheet.Range("A1").Font.Size = 8;
            workSheet.Range("A1").Font.Bold = true;
            workSheet.Range("A1").Font.Italic = true;
            workSheet.Range("A1").Font.Underline = true;
            workSheet.Range("A1").Font.Color = Color.Violet.ToArgb();

            workSheet.Range("A3").Value = "Times New Roman Size:10";
            workSheet.Range("A3").Font.Name = "Times New Roman";
            workSheet.Range("A3").Font.Size = 10;
            workSheet.Range("A3").Font.Color = Color.Orange.ToArgb();

            workSheet.Range("A5").Value = "Comic Sans MS Size:12 WrapText";
            workSheet.Range("A5").Font.Name = "Comic Sans MS";
            workSheet.Range("A5").Font.Size = 12;
            workSheet.Range("A5").WrapText = true;
            workSheet.Range("A5").Font.Color = Color.Navy.ToArgb();

            // HorizontalAlignment
            workSheet.Range("A7").Value = "xlHAlignLeft";
            workSheet.Range("A7").HorizontalAlignment = XlHAlign.xlHAlignLeft;

            workSheet.Range("B7").Value = "xlHAlignCenter";
            workSheet.Range("B7").HorizontalAlignment = XlHAlign.xlHAlignCenter;

            workSheet.Range("C7").Value = "xlHAlignRight";
            workSheet.Range("C7").HorizontalAlignment = XlHAlign.xlHAlignRight;

            workSheet.Range("D7").Value = "xlHAlignJustify";
            workSheet.Range("D7").HorizontalAlignment = XlHAlign.xlHAlignJustify;

            workSheet.Range("E7").Value = "xlHAlignDistributed";
            workSheet.Range("E7").HorizontalAlignment = XlHAlign.xlHAlignDistributed;

            // VerticalAlignment
            workSheet.Range("A9").Value = "xlVAlignTop";
            workSheet.Range("A9").VerticalAlignment = XlVAlign.xlVAlignTop;

            workSheet.Range("B9").Value = "xlVAlignCenter";
            workSheet.Range("B9").VerticalAlignment = XlVAlign.xlVAlignCenter;

            workSheet.Range("C9").Value = "xlVAlignBottom";
            workSheet.Range("C9").VerticalAlignment = XlVAlign.xlVAlignBottom;

            workSheet.Range("D9").Value = "xlVAlignDistributed";
            workSheet.Range("D9").VerticalAlignment = XlVAlign.xlVAlignDistributed;

            workSheet.Range("E9").Value = "xlVAlignJustify";
            workSheet.Range("E9").VerticalAlignment = XlVAlign.xlVAlignJustify;

            // setup rows and columns
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].ColumnWidth = 25;
            workSheet.Columns[3].ColumnWidth = 25;
            workSheet.Columns[4].ColumnWidth = 25;
            workSheet.Columns[5].ColumnWidth = 25;
            workSheet.Rows[9].RowHeight = 25;

            // save the book - utils want build the filename for us
            string workbookFile = utils.File.Combine(HostApplication.RootDirectory, "Example02", DocumentFormat.Normal);
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
            get { return "Example02"; }
        }

        public string Description
        {
            get { return "Font Attributes and Alignment for Cells"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
