using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Tests.Core;
using Core = NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace ExcelTestsCSharp
{
    public class Test02 : ITestPackage 
    {
        #region ITestPackage Member
         
        #endregion

        public string Name
        {
            get { return "Test02"; }
        }

        public string Description
        {
            get { return "Set alignment and font style."; }
        }

        public string OfficeProduct
        {
            get { return "Excel"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Excel.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                // start excel and turn off msg boxes
                application = new Excel.Application();
                application.DisplayAlerts = false;

                // add a new workbook
                Excel.Workbook workBook = application.Workbooks.Add();
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

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    application.Quit();
                    application.Dispose();
                }
            }
        }
    }
}
