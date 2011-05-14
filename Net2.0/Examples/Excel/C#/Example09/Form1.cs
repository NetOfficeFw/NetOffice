using System;
using System.Windows.Forms;
using System.Reflection;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace Example09
{
    public partial class Form1 : Form
    {
        Excel.Application _excelApplication;

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

            Office.CommandBar commandBar;
            Office.CommandBarButton commandBarBtn;
 
            // start excel and turn off msg boxes
            _excelApplication = new Excel.Application();
            _excelApplication.DisplayAlerts = false;
         
            // add a new workbook
            Excel.Workbook workBook = _excelApplication.Workbooks.Add();
             
           
            // add a commandbar popup
            Office.CommandBarPopup  commandBarPopup = (Office.CommandBarPopup)_excelApplication.CommandBars.get_Item("Worksheet Menu Bar").Controls.Add(
                                                                                MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarPopup.Caption = "commandBarPopup";
             
             
            #region few words, how to access the picture
            /*
             you can see we use an own icon via .PasteFace()
             is not possible from outside process boundaries to use the PictureProperty directly
             the reason for is IPictureDisp: http://support.microsoft.com/kb/286460/de
             its not important is early or late binding or managed or unmanaged, the behaviour is always the same
             For example, a COMAddin running as InProcServer and can access the Picture Property
             Use the IconConverter.cs class from this project to convert a image to IPictureDisp
            */
            #endregion

            #region CommandBarButton

            // add a button to the popup
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            Clipboard.SetDataObject(this.Icon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            #endregion

            #region Create a new toolbar

            // add a new toolbar
            commandBar = _excelApplication.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            commandBarBtn.FaceId = 3;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);
            
            // add a dropdown box to the toolbar
            commandBarPopup = (Office.CommandBarPopup)commandBar.Controls.Add(MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            commandBarPopup.Caption = "commandBarPopup";

            // add a button to the popup, we use an own icon for the button
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            Clipboard.SetDataObject(this.Icon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);
            
            #endregion
            
            #region Create a new ContextMenu

            // add a commandbar popup
            commandBarPopup = (Office.CommandBarPopup)_excelApplication.CommandBars.get_Item("Cell").Controls.Add(
                                                            MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarPopup.Caption = "commandBarPopup";

            // add a button to the popup
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
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
            button1.Enabled = false;
            button2.Enabled = true;
        }
  
        private void button2_Click(object sender, EventArgs e)
        {
            _excelApplication.Quit();
            _excelApplication.Dispose();

            button1.Enabled = true;
            button2.Enabled = false;         
        }

        void commandBarBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Click called." });
            Ctrl.Dispose();
        }

        private void UpdateTextbox(string Message)
        {
            textBoxEvents.AppendText(Message + "\r\n");
        }
    }
}
