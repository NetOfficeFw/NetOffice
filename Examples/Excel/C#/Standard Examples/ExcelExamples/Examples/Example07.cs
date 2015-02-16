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
using VB = NetOffice.VBIDEApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.VBIDEApi.Enums;

namespace ExcelExamplesCS4
{
    /// <summary>
    /// Example 7 - Attach VBA Code to a workbook
    /// </summary>
    internal class Example07 : IExample
    {
        #region IExample Member

        public void RunExample()
        {
            bool isFailed = false;
            string workbookFile = null;
            Excel.Application excelApplication = null;
            try
            {           
                // start excel and turn off msg boxes
                excelApplication = new Excel.Application();
                excelApplication.DisplayAlerts = false;
                excelApplication.Visible = false;

                // add a new workbook
                Excel.Workbook workBook = excelApplication.Workbooks.Add();

                // add new global Code Module
                VB.VBComponent globalModule = workBook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                globalModule.Name = "MyNewCodeModule";

                // add a new procedure to the modul
                globalModule.CodeModule.InsertLines(1, "Public Sub HelloWorld(Param as string)\r\n MsgBox \"Hello from NetOffice!\" & vbnewline & Param\r\nEnd Sub");

                // create a click event trigger for the first worksheet
                int linePosition = workBook.VBProject.VBComponents[2].CodeModule.CreateEventProc("BeforeDoubleClick", "Worksheet");
                workBook.VBProject.VBComponents[2].CodeModule.InsertLines(linePosition + 1, "HelloWorld \"BeforeDoubleClick\"");

                // display info in the worksheet
                Excel.Worksheet sheet = (Excel.Worksheet)workBook.Worksheets[1];

                sheet.Cells[2, 2].Value = "This workbook contains dynamic created VBA Moduls and Event Code";
                sheet.Cells[5, 2].Value = "Open the VBA Editor to see the code";
                sheet.Cells[8, 2].Value = "Do a double click to catch the BeforeDoubleClick Event from this Worksheet.";

                // save the book 
                string fileExtension = GetDefaultExtension(excelApplication);
                XlFileFormat fileFormat = GetFileFormat(excelApplication);
                workbookFile = string.Format("{0}\\Example07{1}", HostApplication.RootDirectory, fileExtension);
                workBook.SaveAs(workbookFile, fileFormat);

            }
            catch (System.Runtime.InteropServices.COMException throwedException)
            {
                isFailed = true;
                HostApplication.ShowErrorDialog("VBA Error", throwedException);
            }
            finally
            {
                // close excel and dispose reference
                excelApplication.Quit();
                excelApplication.Dispose();

                if ((null != workbookFile) && (!isFailed))
                    HostApplication.ShowFinishDialog(null, workbookFile);
            }
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return HostApplication.LCID == 1033 ? "Example07" : "Beispiel07"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Attach VBA Code to a workbook. The option 'Trust Visual Basic projects' must be set." : "Dynamisches hinzufügen von VBA Code zu einem Workbook. Die Option 'Visual Basic Projekten vertrauen' muss aktiviert sein."; }
        }

        public UserControl Panel
        {
            get { return null; }
        }
  
        #endregion

        #region Properties

        /// <summary>
        /// Current Example Host
        /// </summary>
        internal IHost HostApplication { get; private set; }

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
        /// Returns the valid file extension for the instance. for example ".xls" or ".xlsx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Excel.Application application)
        {
            double Version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (Version >= 12.00)
                return ".xlsm";
            else
                return ".xls";
        }

        /// <summary>
        /// Returns the valid file format for the instance. Documents with macro's need a bit xtra config since 2007
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the format</returns>
        private static XlFileFormat GetFileFormat(Excel.Application application)
        {
            double Version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (Version >= 12.00)
                return XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
            else
                return XlFileFormat.xlExcel7;
        }

        #endregion
    }
}
