using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace Example08
{
    public partial class Form1 : Form
    {
        
        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Form1()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }
     
        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // we enable the event support
            LateBindingApi.Core.Settings.EnableEvents = true;
 
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            excelApplication.Visible = true;

            /*
            we register some events. note: the event trigger was called from excel, means an other Thread
            remove the Quit() call below and check out more events if you want
            you can get event notifys from various objects: Application or Workbook or Worksheet for example
            */

            excelApplication.NewWorkbookEvent += new Excel.Application_NewWorkbookEventHandler(excelApplication_NewWorkbook);
            excelApplication.WorkbookBeforeCloseEvent += new Excel.Application_WorkbookBeforeCloseEventHandler(excelApplication_WorkbookBeforeClose);
            excelApplication.WorkbookActivateEvent += new Excel.Application_WorkbookActivateEventHandler(excelApplication_WorkbookActivate);
            excelApplication.WorkbookDeactivateEvent += new Excel.Application_WorkbookDeactivateEventHandler(excelApplication_WorkbookDeactivate);
            excelApplication.SheetActivateEvent += new Excel.Application_SheetActivateEventHandler(_excelApplication_SheetActivateEvent);
            excelApplication.SheetDeactivateEvent += new Excel.Application_SheetDeactivateEventHandler(_excelApplication_SheetDeactivateEvent);

            // add a new workbook add a sheet and close
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            workBook.Worksheets.Add();
            workBook.Close();

            excelApplication.Quit();
            excelApplication.Dispose();
        }

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
    }
}
    