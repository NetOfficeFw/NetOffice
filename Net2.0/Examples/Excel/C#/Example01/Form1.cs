using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.VBIDEApi.Enums;

namespace Example01
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];


            // draw back color and perform the BorderAround method
            workSheet.get_Range("$B2:$B5").Interior.Color = ToDouble(Color.DarkGreen);
            workSheet.get_Range("$B2:$B5").BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic);
            
            // draw back color and border the range explicitly
            workSheet.get_Range("$D2:$D5").Interior.Color = ToDouble(Color.DarkGreen);
            workSheet.get_Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlDouble;
            workSheet.get_Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].Weight = 4;
            workSheet.get_Range("$D2:$D5").Borders[XlBordersIndex.xlInsideHorizontal].Color = ToDouble(Color.Black);


            // save the book 
            string fileExtension = GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example01{1}", Application.StartupPath, fileExtension);
            workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }

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
            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".xlsx";
            else
                return ".xls";
        }
        
        #endregion

    }
}
