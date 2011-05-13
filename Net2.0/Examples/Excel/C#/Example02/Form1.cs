using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums; 
 
namespace Example02
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];
           
            // font action
            workSheet.get_Range("A1").Value = "Arial Size:8 Bold Italic Underline";
            workSheet.get_Range("A1").Font.Name = "Arial";
            workSheet.get_Range("A1").Font.Size = 8;
            workSheet.get_Range("A1").Font.Bold = true;
            workSheet.get_Range("A1").Font.Italic = true;
            workSheet.get_Range("A1").Font.Underline = true;
            workSheet.get_Range("A1").Font.Color = Color.Violet.ToArgb();

            workSheet.get_Range("A3").Value = "Times New Roman Size:10";
            workSheet.get_Range("A3").Font.Name = "Times New Roman";
            workSheet.get_Range("A3").Font.Size = 10;
            workSheet.get_Range("A3").Font.Color = Color.Orange.ToArgb();

            workSheet.get_Range("A5").Value = "Comic Sans MS Size:12 WrapText";
            workSheet.get_Range("A5").Font.Name = "Comic Sans MS";
            workSheet.get_Range("A5").Font.Size = 12;
            workSheet.get_Range("A5").WrapText = true;
            workSheet.get_Range("A5").Font.Color = Color.Navy.ToArgb();
            
            // HorizontalAlignment
            workSheet.get_Range("A7").Value = "xlHAlignLeft";
            workSheet.get_Range("A7").HorizontalAlignment = XlHAlign.xlHAlignLeft;

            workSheet.get_Range("B7").Value = "xlHAlignCenter";
            workSheet.get_Range("B7").HorizontalAlignment  = XlHAlign.xlHAlignCenter;

            workSheet.get_Range("C7").Value = "xlHAlignRight";
            workSheet.get_Range("C7").HorizontalAlignment = XlHAlign.xlHAlignRight;
          
            workSheet.get_Range("D7").Value = "xlHAlignJustify";
            workSheet.get_Range("D7").HorizontalAlignment = XlHAlign.xlHAlignJustify;

            workSheet.get_Range("E7").Value = "xlHAlignDistributed";
            workSheet.get_Range("E7").HorizontalAlignment = XlHAlign.xlHAlignDistributed;

            // VerticalAlignment
            workSheet.get_Range("A9").Value = "xlVAlignTop";
            workSheet.get_Range("A9").VerticalAlignment = XlVAlign.xlVAlignTop;

            workSheet.get_Range("B9").Value = "xlVAlignCenter";
            workSheet.get_Range("B9").VerticalAlignment = XlVAlign.xlVAlignCenter;

            workSheet.get_Range("C9").Value = "xlVAlignBottom";
            workSheet.get_Range("C9").VerticalAlignment = XlVAlign.xlVAlignBottom;

            workSheet.get_Range("D9").Value = "xlVAlignDistributed";
            workSheet.get_Range("D9").VerticalAlignment = XlVAlign.xlVAlignDistributed;

            workSheet.get_Range("E9").Value = "xlVAlignJustify";
            workSheet.get_Range("E9").VerticalAlignment = XlVAlign.xlVAlignJustify;
            
            // setup rows and columns
            workSheet.Columns[1, Missing.Value].AutoFit();
            workSheet.Columns[2, Missing.Value].ColumnWidth = 25;
            workSheet.Columns[3, Missing.Value].ColumnWidth = 25;
            workSheet.Columns[4, Missing.Value].ColumnWidth = 25;
            workSheet.Columns[5, Missing.Value].ColumnWidth = 25;
            workSheet.Rows[9, Missing.Value].RowHeight = 25;

            // save the book 
            string fileExtension = GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example02{1}", Environment.CurrentDirectory, fileExtension);
            workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }

        #region Helper

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
