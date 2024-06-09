using System;
using System.Windows.Forms;
using System.Globalization;
using ExampleBase;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools.Contribution;

namespace ExcelExamplesCS4
{
    /// <summary>
    /// Example 3 - Using Numberformats
    /// </summary>
    internal class Example03 : IExample
    {
        public void RunExample()
        {
            // start excel and turn Application msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // create a utils instance, not need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(excelApplication);

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

            // the given thread culture in all NetOffice calls are stored in NetOffice.Settings.
            // you can change the culture of course. Default is en-us.
            CultureInfo cultureInfo = NetOffice.Settings.Default.ThreadCulture;
            string Pattern1 = string.Format("0{0}00", cultureInfo.NumberFormat.CurrencyDecimalSeparator);
            string Pattern2 = string.Format("#{1}##0{0}00", cultureInfo.NumberFormat.CurrencyDecimalSeparator, cultureInfo.NumberFormat.CurrencyGroupSeparator);

            workSheet.Range("A1").Value = "Type";
            workSheet.Range("B1").Value = "Value";
            workSheet.Range("C1").Value = "Formatted " + Pattern1;
            workSheet.Range("D1").Value = "Formatted " + Pattern2;

            int integerValue = 532234;
            workSheet.Range("A3").Value = "Integer";
            workSheet.Range("B3").Value = integerValue;
            workSheet.Range("C3").Value = integerValue;
            workSheet.Range("C3").NumberFormat = Pattern1;
            workSheet.Range("D3").Value = integerValue;
            workSheet.Range("D3").NumberFormat = Pattern2;

            double doubleValue = 23172.64;
            workSheet.Range("A4").Value = "double";
            workSheet.Range("B4").Value = doubleValue;
            workSheet.Range("C4").Value = doubleValue;
            workSheet.Range("C4").NumberFormat = Pattern1;
            workSheet.Range("D4").Value = doubleValue;
            workSheet.Range("D4").NumberFormat = Pattern2;

            float floatValue = 84345.9132f;
            workSheet.Range("A5").Value = "float";
            workSheet.Range("B5").Value = floatValue;
            workSheet.Range("C5").Value = floatValue;
            workSheet.Range("C5").NumberFormat = Pattern1;
            workSheet.Range("D5").Value = floatValue;
            workSheet.Range("D5").NumberFormat = Pattern2;

            Decimal decimalValue = 7251231.313367m;
            workSheet.Range("A6").Value = "Decimal";
            workSheet.Range("B6").Value = decimalValue;
            workSheet.Range("C6").Value = decimalValue;
            workSheet.Range("C6").NumberFormat = Pattern1;
            workSheet.Range("D6").Value = decimalValue;
            workSheet.Range("D6").NumberFormat = Pattern2;

            workSheet.Range("A9").Value = "DateTime";
            workSheet.Range("B10").Value = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.FullDateTimePattern;
            workSheet.Range("C10").Value = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.LongDatePattern;
            workSheet.Range("D10").Value = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.ShortDatePattern;
            workSheet.Range("E10").Value = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.LongTimePattern;
            workSheet.Range("F10").Value = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.ShortTimePattern;

            // DateTime
            DateTime dateTimeValue = DateTime.Now;
            workSheet.Range("B11").Value = dateTimeValue;
            workSheet.Range("B11").NumberFormat = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.FullDateTimePattern;

            workSheet.Range("C11").Value = dateTimeValue;
            workSheet.Range("C11").NumberFormat = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.LongDatePattern;

            workSheet.Range("D11").Value = dateTimeValue;
            workSheet.Range("D11").NumberFormat = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.ShortDatePattern;

            workSheet.Range("E11").Value = dateTimeValue;
            workSheet.Range("E11").NumberFormat = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.LongTimePattern;

            workSheet.Range("F11").Value = dateTimeValue;
            workSheet.Range("F11").NumberFormat = NetOffice.Settings.Default.ThreadCulture.DateTimeFormat.ShortTimePattern;

            // string
            workSheet.Range("A14").Value = "String";
            workSheet.Range("B14").Value = "This is a sample String";
            workSheet.Range("B14").NumberFormat = "@";

            // number as string
            workSheet.Range("B15").Value = "513";
            workSheet.Range("B15").NumberFormat = "@";

            // set colums
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();

            // save the book 
            string workbookFile = utils.File.Combine(HostApplication.RootDirectory, "Example03", DocumentFormat.Normal);
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
            get { return "Example03"; }
        }

        public string Description
        {
            get { return "Using Numberformats"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
