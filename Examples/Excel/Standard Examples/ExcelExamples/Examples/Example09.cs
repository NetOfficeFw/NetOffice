using System;
using System.Windows.Forms;
using ExampleBase;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using NetOffice.OfficeApi.Enums;

namespace ExcelExamplesCS4
{
    /// <summary>
    /// Example 9 - Customize classic UI and recieve events
    /// </summary>
    partial class Example09 : UserControl , IExample
    {
        #region Fields/Delegates

        private Excel.Application _excelApplication;
        private delegate void UpdateEventTextDelegate(string Message);
        private UpdateEventTextDelegate _updateDelegate;

        #endregion

        #region Ctor

        public Example09()
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
            get { return "Example09"; }
        }

        public string Description
        {
            get { return "Customize classic UI without ribbons and recieve click events"; }
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

        #region UI Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {        
            // start excel and turn off msg boxes
            _excelApplication = new Excel.Application();
            _excelApplication.DisplayAlerts = false;

            Office.CommandBar commandBar = null;
            Office.CommandBarButton commandBarBtn = null;

            // add a new workbook
            Excel.Workbook workBook = _excelApplication.Workbooks.Add();

            //var test = _excelApplication.CommandBars["Worksheet Menu Bar"].Controls.Add(
            //        MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);

            // add a commandbar popup
            Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)
                _excelApplication.CommandBars["Worksheet Menu Bar"].Controls.Add(
                    MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarPopup.Caption = "commandBarPopup";

            #region few words, how to access the picture
            /*
             you can see we use an own icon via .PasteFace()
             is not possible from outside process boundaries to use the PictureProperty directly
             the reason for is IPictureDisp: http://support.microsoft.com/kb/286460/de
             its not important is early or late binding or managed or unmanaged, the behaviour is always the same
             For example, a COMAddin running as InProcServer and can access the Picture Property
            */
            #endregion

            #region CommandBarButton

            // add a button to the popup
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            Clipboard.SetDataObject(HostApplication.DisplayIcon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            #endregion

            #region Create a new toolbar

            // add a new toolbar
            commandBar = _excelApplication.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            commandBarBtn.FaceId = 3;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            // add a dropdown box to the toolbar
            commandBarPopup = (Office.CommandBarPopup)commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarPopup.Caption = "commandBarPopup";

            // add a button to the popup, we use an own icon for the button
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            Clipboard.SetDataObject(HostApplication.DisplayIcon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            #endregion

            #region Create a new ContextMenu

            // add a commandbar popup
            commandBarPopup = (Office.CommandBarPopup)_excelApplication.CommandBars["Cell"].Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarPopup.Caption = "commandBarPopup";

            // add a button to the popup
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            commandBarBtn.FaceId = 9;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            #endregion

            #region Display info

            Excel.Worksheet sheet = (Excel.Worksheet)workBook.Worksheets[1];
            sheet.Cells[2, 2].Value = "this excel instance contains 3 custom menus";
            sheet.Cells[3, 2].Value = "the main menu, the toolbar menu and the cell context menu";
            sheet.Cells[4, 2].Value = "in this case the menus are temporaily created";
            sheet.Cells[5, 2].Value = "they are not persistant and needs no unload event or something like this";
            sheet.Cells[6, 2].Value = "you can also create persistant menus if you want";

            #endregion

            // make visible & set buttons
            _excelApplication.Visible = true;
            buttonStartExample.Enabled = false;
            buttonQuitExample.Enabled = true;
        }

        private void buttonQuitExample_Click(object sender, EventArgs e)
        {
            _excelApplication.Quit();
            _excelApplication.Dispose();

            buttonStartExample.Enabled = true;
            buttonQuitExample.Enabled = false;
        }

        #endregion

        #region Excel Trigger

        private void commandBarBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Click called." });
            Ctrl.Dispose();
        }

        private void UpdateTextbox(string Message)
        {
            textBoxEvents.AppendText(Message + "\r\n");
        }

        #endregion
    }
}
