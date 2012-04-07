using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExampleBase;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace ExcelExamples
{
    partial class Example08 : UserControl , IExample
    {
        IHost _hostApplication;

        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Example08()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }
         
        #region IExample Member

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example08" : "Beispiel08"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Using Events" : "Verwenden von Ereignissen"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }
 
        #endregion

        #region UI Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {            
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            excelApplication.Visible = true;

            // we register some events. note: the event trigger was called from excel, means another Thread
            // you can get event notifys from various objects: Application or Workbook or Worksheet for example
            excelApplication.NewWorkbookEvent += new Excel.Application_NewWorkbookEventHandler(excelApplication_NewWorkbook);
            excelApplication.WorkbookBeforeCloseEvent += new Excel.Application_WorkbookBeforeCloseEventHandler(excelApplication_WorkbookBeforeClose);
            excelApplication.WorkbookActivateEvent += new Excel.Application_WorkbookActivateEventHandler(excelApplication_WorkbookActivate);
            excelApplication.WorkbookDeactivateEvent += new Excel.Application_WorkbookDeactivateEventHandler(excelApplication_WorkbookDeactivate);
            excelApplication.SheetActivateEvent += new Excel.Application_SheetActivateEventHandler(_excelApplication_SheetActivateEvent);
            excelApplication.SheetDeactivateEvent += new Excel.Application_SheetDeactivateEventHandler(_excelApplication_SheetDeactivateEvent);

            // add a new workbook, add a sheet and close
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            workBook.Worksheets.Add();
            workBook.Close();

            excelApplication.Quit();
            excelApplication.Dispose();
        }
        
        #endregion

        #region Excel Trigger

        void _excelApplication_SheetDeactivateEvent(COMObject Sh)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event SheetDeactivate called." });
            Sh.Dispose();
        }

        void _excelApplication_SheetActivateEvent(COMObject Sh)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event SheetActivate called." });
            Sh.Dispose();
        }

        void excelApplication_NewWorkbook(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event NewWorkbook called." });
            Wb.Dispose();
        }

        void excelApplication_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookBeforeClose called." });
            Wb.Dispose();
        }

        void excelApplication_WorkbookActivate(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookActivate called." });
            Wb.Dispose();
        }

        void excelApplication_WorkbookDeactivate(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookDeactivate called." });
            Wb.Dispose();
        }

        private void UpdateTextbox(string message)
        {
            textBoxEvents.AppendText(message + "\r\n");
        }

        #endregion
    }
}
