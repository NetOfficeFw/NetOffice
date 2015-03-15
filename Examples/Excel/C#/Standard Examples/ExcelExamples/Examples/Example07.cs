using System;
using System.Windows.Forms;
using System.Globalization;
using ExampleBase;
using Office = NetOffice.OfficeApi;
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

                // create a utils instance, not need for but helpful to keep the lines of code low
                Excel.Tools.CommonUtils utils = new Excel.Tools.CommonUtils(excelApplication);

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
                XlFileFormat fileFormat = GetFileFormat(excelApplication);
                workbookFile = utils.File.Combine(HostApplication.RootDirectory, "Example07", Excel.Tools.DocumentFormat.Macros);
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
