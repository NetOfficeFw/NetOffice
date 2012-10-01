using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using NetOffice;
using Core = NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace ExcelTestsCSharp
{
    /// <summary>
    /// events
    /// </summary>
    public class Test06 : ITestPackage
    {
        bool _sheetDeactivateEvent;
        bool _sheetActivateEvent;
        bool _newWorkbookEvent;
        bool _workbookBeforeCloseEvent;
        bool _workbookActivateEvent;
        bool _workbookDeactivateEvent;

        #region TestPackage Member

        public string Name
        {
            get { return "Test06"; }
        }

        public string Description
        {
            get { return "Using events"; }
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
                application.Visible = true;

                application.NewWorkbookEvent += new Excel.Application_NewWorkbookEventHandler(ExcelApplication_NewWorkbook);
                application.WorkbookBeforeCloseEvent += new Excel.Application_WorkbookBeforeCloseEventHandler(ExcelApplication_WorkbookBeforeClose);
                application.WorkbookActivateEvent += new Excel.Application_WorkbookActivateEventHandler(ExcelApplication_WorkbookActivate);
                application.WorkbookDeactivateEvent += new Excel.Application_WorkbookDeactivateEventHandler(ExcelApplication_WorkbookDeactivate);
                application.SheetActivateEvent += new Excel.Application_SheetActivateEventHandler(ExcelApplication_SheetActivateEvent);
                application.SheetDeactivateEvent += new Excel.Application_SheetDeactivateEventHandler(ExcelApplication_SheetDeactivateEvent);

                // add a new workbook add a sheet and close
                Excel.Workbook workBook = application.Workbooks.Add();
                workBook.Worksheets.Add();
                workBook.Close();

                if (_newWorkbookEvent && _workbookBeforeCloseEvent && _sheetActivateEvent && _sheetDeactivateEvent && _workbookActivateEvent && _workbookDeactivateEvent)
                    return new TestResult(true, DateTime.Now.Subtract(startTime), "",  null, "");
                else
                {
                    string errorMessage = "";
                    if (!_newWorkbookEvent)
                        errorMessage += "NewWorkbookEvent failed ";
                    if (!_workbookBeforeCloseEvent)
                        errorMessage += "WorkbookBeforeCloseEvent failed ";
                    if (!_sheetActivateEvent)
                        errorMessage += "WorkbookActivateEvent failed ";
                    if (!_sheetDeactivateEvent)
                        errorMessage += "WorkbookDeactivateEvent failed ";
                    if (!_workbookActivateEvent)
                        errorMessage += "SheetActivateEvent failed ";
                    if (!_workbookDeactivateEvent)
                        errorMessage += "SheetDeactivateEvent failed ";

                    return new TestResult(true, DateTime.Now.Subtract(startTime), errorMessage, null, "");
                }
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

        #endregion

        void ExcelApplication_SheetDeactivateEvent(COMObject Sh)
        {
            _sheetDeactivateEvent = true;
            Sh.Dispose();
        }

        void ExcelApplication_SheetActivateEvent(COMObject Sh)
        {
            _sheetActivateEvent = true;
            Sh.Dispose();
        }

        void ExcelApplication_NewWorkbook(Excel.Workbook Wb)
        {
            _newWorkbookEvent = true;
            Wb.Dispose();
        }

        void ExcelApplication_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            _workbookBeforeCloseEvent = true;
            Wb.Dispose();
        }

        void ExcelApplication_WorkbookActivate(Excel.Workbook Wb)
        {
            _workbookActivateEvent = true;
            Wb.Dispose();
        }

        void ExcelApplication_WorkbookDeactivate(Excel.Workbook Wb)
        {
            _workbookDeactivateEvent = true;
            Wb.Dispose();
        }
    }
}
