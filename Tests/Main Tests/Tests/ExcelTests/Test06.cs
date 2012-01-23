using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using LateBindingApi.Core;
using Core = LateBindingApi.Core;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace ExcelTests
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

        public bool DoTest(string logFilePath)
        {
            Core.DebugConsole.FileName = System.IO.Path.Combine(logFilePath, "ExcelTests.Test06.log");
            Core.DebugConsole.AppendTimeInfoEnabled = true;
            Core.DebugConsole.Mode = LateBindingApi.Core.ConsoleMode.LogFile;

            Excel.Application application = null;
            try
            {
                // Initialize NetOffice
                LateBindingApi.Core.Factory.Initialize();

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

                if (!_newWorkbookEvent)
                    Core.DebugConsole.WriteLine("NewWorkbookEvent not called");

                if (!_workbookBeforeCloseEvent)
                    Core.DebugConsole.WriteLine("WorkbookBeforeCloseEvent not called");
               
                if (!_sheetActivateEvent)
                    Core.DebugConsole.WriteLine("WorkbookActivateEvent not called");
 
                if (!_sheetDeactivateEvent)
                    Core.DebugConsole.WriteLine("WorkbookDeactivateEvent not called");

                if (!_workbookActivateEvent)
                    Core.DebugConsole.WriteLine("SheetActivateEvent not called");

                if (!_workbookDeactivateEvent)
                    Core.DebugConsole.WriteLine("SheetDeactivateEvent not called");

                if (_newWorkbookEvent && _workbookBeforeCloseEvent && _sheetActivateEvent
                    && _sheetDeactivateEvent && _workbookActivateEvent && _workbookDeactivateEvent)
                    return true;
                else
                    return false;
            }
            catch (Exception exception)
            {
                string message = exception.Message;
                Console.WriteLine("An error occured{1}{1}{0}", message, Environment.NewLine);
                return false;
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
