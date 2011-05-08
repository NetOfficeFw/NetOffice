using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = LateBindingApi.ExcelApi;
using VBE = LateBindingApi.VBIDEApi;
using LateBindingApi.VBIDEApi.Enums;
    
namespace Example7
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
                VBE.VBComponent globalModule = workBook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                globalModule.Name = "MyNewCodeModule";

                // add a new procedure to the modul
                globalModule.CodeModule.InsertLines(1, "Public Sub HelloWorld(Param as string)\r\n MsgBox \"Hello World!\" & vbnewline & Param\r\nEnd Sub");
                 
                // create a click event trigger for the first worksheet
                int linePosition = workBook.VBProject.VBComponents.Item(2).CodeModule.CreateEventProc("BeforeDoubleClick", "Worksheet");
                workBook.VBProject.VBComponents.Item(2).CodeModule.InsertLines(linePosition + 1, "HelloWorld \"BeforeDoubleClick\"");

                // display info in the worksheet
                Excel.Worksheet sheet = (Excel.Worksheet)workBook.Worksheets[1];

                sheet.Cells[2, 2].Value = "This workbook contains dynamic created VBA Moduls and Event Code";
                sheet.Cells[5, 2].Value = "Open the VBA Editor to see the code";
                sheet.Cells[8, 2].Value = "Do a double click to catch the BeforeDoubleClick Event from this Worksheet.";          
               
                // save the book 
                string fileExtension = GetDefaultExtension(excelApplication);
                workbookFile = string.Format("{0}\\Example07{1}", Environment.CurrentDirectory, fileExtension);
                workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, LateBindingApi.ExcelApi.Enums.XlSaveAsAccessMode.xlExclusive);
            }
            catch (System.Reflection.TargetInvocationException throwedException)
            {
                string message = string.Format("An error is occured.{0}ExceptionTrace:{0}", Environment.NewLine);
               
                Exception exception = throwedException;
                while (null != exception)
                {
                    message += string.Format("{0}{1}", exception.Message, Environment.NewLine); 
                    exception = exception.InnerException;
                }

                MessageBox.Show(message); 
            }
            finally
            {
                // close excel and dispose reference
                excelApplication.Quit();
                excelApplication.Dispose();

                if (null != workbookFile)
                { 
                    FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
                    fDialog.ShowDialog(this);
                }
            }
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
