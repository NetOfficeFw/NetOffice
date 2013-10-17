using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using Tests.Core;
using Core = NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace ExcelTestsCSharp
{
    /// <summary>
    /// numberformat in cells
    /// </summary>
    public class Test03 : ITestPackage 
    {
        #region ITestPackage Member
         
        public string Name
        {
            get { return "Test03"; }
        }

        public string Description
        {
            get { return "Set numberformat in cells."; }
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
                // start excel and turn off Application msg boxes
                application = new Excel.Application();
                application.DisplayAlerts = false;

                // add a new workbook
                Excel.Workbook workBook = application.Workbooks.Add();
                Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

                /* some kind of numerics */

                // the given thread culture in all latebinding calls are stored in NetOffice.Settings.
                // you can change the culture. default is en-us.
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
                workSheet.get_Range("D4").NumberFormat = Pattern2;

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
                workSheet.Range("B10").Value = cultureInfo.DateTimeFormat.FullDateTimePattern;
                workSheet.Range("C10").Value = cultureInfo.DateTimeFormat.LongDatePattern;
                workSheet.Range("D10").Value = cultureInfo.DateTimeFormat.ShortDatePattern;
                workSheet.Range("E10").Value = cultureInfo.DateTimeFormat.LongTimePattern;
                workSheet.Range("F10").Value = cultureInfo.DateTimeFormat.ShortTimePattern;

                // DateTime
                DateTime dateTimeValue = DateTime.Now;
                workSheet.Range("B11").Value = dateTimeValue;
                workSheet.Range("B11").NumberFormat = cultureInfo.DateTimeFormat.FullDateTimePattern;

                workSheet.Range("C11").Value = dateTimeValue;
                workSheet.Range("C11").NumberFormat = cultureInfo.DateTimeFormat.LongDatePattern;

                workSheet.Range("D11").Value = dateTimeValue;
                workSheet.Range("D11").NumberFormat = cultureInfo.DateTimeFormat.ShortDatePattern;

                workSheet.Range("E11").Value = dateTimeValue;
                workSheet.Range("E11").NumberFormat = cultureInfo.DateTimeFormat.LongTimePattern;

                workSheet.Range("F11").Value = dateTimeValue;
                workSheet.Range("F11").NumberFormat = cultureInfo.DateTimeFormat.ShortTimePattern;

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

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception,"");
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

        #endregion
    }
}
