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
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

            /*do background color for cells*/

            string listSeperator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;

            // draw the face
            string rangeAdressFace = string.Format("$C10:$M10{0}$C30:$M30{0}$C11:$C30{0}$M11:$M30", listSeperator);
            workSheet.get_Range(rangeAdressFace).Interior.Color = ToDouble(Color.DarkGreen);

            string rangeAdressEyes = string.Format("$F14{0}$J14", listSeperator);
            workSheet.get_Range(rangeAdressEyes).Interior.Color = ToDouble(Color.Black);

            string rangeAdressNoise = string.Format("$G18:$I19", listSeperator);
            workSheet.get_Range(rangeAdressNoise).Interior.Color = ToDouble(Color.DarkGreen);

            string rangeAdressMouth = string.Format("$F26{0}$J26{0}$G27:$I27", listSeperator);
            workSheet.get_Range(rangeAdressMouth).Interior.Color = ToDouble(Color.DarkGreen);

            // border the face with the border arround method
            workSheet.get_Range(rangeAdressFace).BorderAround(XlLineStyle.xlDashDot, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexNone, Color.BlueViolet.ToArgb());
            workSheet.get_Range(rangeAdressEyes).BorderAround(XlLineStyle.xlDashDot, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexNone, Color.BlueViolet.ToArgb());
            workSheet.get_Range(rangeAdressNoise).BorderAround(XlLineStyle.xlDouble, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexNone, Color.BlueViolet.ToArgb());

            // border explicitly
            workSheet.get_Range(rangeAdressMouth).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            workSheet.get_Range(rangeAdressMouth).Borders[XlBordersIndex.xlEdgeBottom].Weight = 4;
            workSheet.get_Range(rangeAdressMouth).Borders[XlBordersIndex.xlEdgeBottom].Color = ToDouble(Color.Black);

            // save the book 
            string fileExtension = GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example01{1}", Application.StartupPath, fileExtension);
            workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, 
                                                                                        XlSaveAsAccessMode.xlExclusive);

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
