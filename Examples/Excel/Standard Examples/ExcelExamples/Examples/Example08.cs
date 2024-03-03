using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;

namespace ExcelExamplesCS4
{
    /// <summary>
    /// Example 8 - Using Events
    /// </summary>
    partial class Example08 : UserControl , IExample
    {
        #region Fields/Delegates

        private delegate void UpdateEventTextDelegate(string Message);
        private UpdateEventTextDelegate _updateDelegate;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Example08()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }

        #endregion

        #region IExample Member

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return "Example08"; }
        }

        public string Description
        {
            get { return "Using Events"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }
 
        #endregion

        #region Properties

        /// <summary>
        /// Current Example Host
        /// </summary>
        internal IHost HostApplication { get; private set; }

        #endregion

        #region Methods

        private void UpdateTextbox(string message)
        {
            textBoxEvents.AppendText(message + "\r\n");
        }

        #endregion

        #region UI Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {            
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

        private void _excelApplication_SheetDeactivateEvent(ICOMObject sh)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event SheetDeactivate called." });
            sh.Dispose();
        }

        private void _excelApplication_SheetActivateEvent(ICOMObject sh)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event SheetActivate called." });
            sh.Dispose();
        }

        private void excelApplication_NewWorkbook(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event NewWorkbook called." });
            Wb.Dispose();
        }

        private void excelApplication_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookBeforeClose called." });
            Wb.Dispose();
        }

        private void excelApplication_WorkbookActivate(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookActivate called." });
            Wb.Dispose();
        }

        private void excelApplication_WorkbookDeactivate(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookDeactivate called." });
            Wb.Dispose();
        }

        #endregion
    }
}
