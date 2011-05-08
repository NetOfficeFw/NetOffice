using System;
using System.Reflection;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = LateBindingApi.ExcelApi;
using LateBindingApi.ExcelApi.Enums; 
 
namespace Example3
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

            // start excel and turn Application msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];
            
            /* some kind of numerics */
           
            // the given thread culture in all latebinding calls are stored in LateBindingApi.Core.Settings.
            // you can change the culture. default is en-us.
            CultureInfo cultureInfo = LateBindingApi.Core.Settings.ThreadCulture;
            string Pattern1 = string.Format("0{0}00", cultureInfo.NumberFormat.CurrencyDecimalSeparator);
            string Pattern2 = string.Format("#{1}##0{0}00", cultureInfo.NumberFormat.CurrencyDecimalSeparator, cultureInfo.NumberFormat.CurrencyGroupSeparator);
            
            workSheet.get_Range("A1").Value = "Type";
            workSheet.get_Range("B1").Value = "Value";
            workSheet.get_Range("C1").Value = "Formatted " + Pattern1;
            workSheet.get_Range("D1").Value = "Formatted " + Pattern2;

            int integerValue = 532234;
            workSheet.get_Range("A3").Value = "Integer";
            workSheet.get_Range("B3").Value = integerValue;
            workSheet.get_Range("C3").Value = integerValue;
            workSheet.get_Range("C3").NumberFormat = Pattern1;
            workSheet.get_Range("D3").Value = integerValue;
            workSheet.get_Range("D3").NumberFormat = Pattern2;

            double doubleValue = 23172.64;
            workSheet.get_Range("A4").Value = "double";
            workSheet.get_Range("B4").Value = doubleValue;
            workSheet.get_Range("C4").Value = doubleValue;
            workSheet.get_Range("C4").NumberFormat= Pattern1;
            workSheet.get_Range("D4").Value = doubleValue;
            workSheet.get_Range("D4").NumberFormat = Pattern2;

            float floatValue = 84345.9132f;
            workSheet.get_Range("A5").Value = "float";
            workSheet.get_Range("B5").Value = floatValue;
            workSheet.get_Range("C5").Value = floatValue;
            workSheet.get_Range("C5").NumberFormat = Pattern1;
            workSheet.get_Range("D5").Value = floatValue;
            workSheet.get_Range("D5").NumberFormat = Pattern2;

            Decimal decimalValue = 7251231.313367m;
            workSheet.get_Range("A6").Value = "Decimal";
            workSheet.get_Range("B6").Value = decimalValue;
            workSheet.get_Range("C6").Value = decimalValue;
            workSheet.get_Range("C6").NumberFormat = Pattern1;
            workSheet.get_Range("D6").Value = decimalValue;
            workSheet.get_Range("D6").NumberFormat =  Pattern2;

            workSheet.get_Range("A9").Value = "DateTime";
            workSheet.get_Range("B10").Value = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.FullDateTimePattern;
            workSheet.get_Range("C10").Value = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.LongDatePattern;
            workSheet.get_Range("D10").Value = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.ShortDatePattern;
            workSheet.get_Range("E10").Value = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.LongTimePattern;
            workSheet.get_Range("F10").Value = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.ShortTimePattern;

            // DateTime
            DateTime dateTimeValue = DateTime.Now;
            workSheet.get_Range("B11").Value = dateTimeValue;
            workSheet.get_Range("B11").NumberFormat = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.FullDateTimePattern;

            workSheet.get_Range("C11").Value = dateTimeValue;
            workSheet.get_Range("C11").NumberFormat = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.LongDatePattern;

            workSheet.get_Range("D11").Value = dateTimeValue;
            workSheet.get_Range("D11").NumberFormat = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.ShortDatePattern;

            workSheet.get_Range("E11").Value = dateTimeValue;
            workSheet.get_Range("E11").NumberFormat = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.LongTimePattern;

            workSheet.get_Range("F11").Value = dateTimeValue;
            workSheet.get_Range("F11").NumberFormat = LateBindingApi.Core.Settings.ThreadCulture.DateTimeFormat.ShortTimePattern;

            // string
            workSheet.get_Range("A14").Value = "String";
            workSheet.get_Range("B14").Value = "This is a sample String";
            workSheet.get_Range("B14").NumberFormat = "@";

            // number as string
            workSheet.get_Range("B15").Value = "513";
            workSheet.get_Range("B15").NumberFormat = "@";

            // set colums
            workSheet.Columns[1, Missing.Value].AutoFit();
            workSheet.Columns[2, Missing.Value].AutoFit();
            workSheet.Columns[3, Missing.Value].AutoFit();
            workSheet.Columns[4, Missing.Value].AutoFit();
       
            // save the book 
            string fileExtension = GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example03{1}", Environment.CurrentDirectory, fileExtension);
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
