using System;
using System.Reflection; 
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using PowerPoint = NetOffice.PowerPointApi;
using Office = NetOffice.OfficeApi; 
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace Example07
{
    public partial class Form1 : Form
    {
        PowerPoint.Application _powerApplication;

        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Form1()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            Office.CommandBar commandBar;
            Office.CommandBarButton commandBarBtn;

            // start powerpoint
            _powerApplication = new PowerPoint.Application();
            
            // add a new presentation with one new slide
            PowerPoint.Presentation presentation = _powerApplication.Presentations.Add(MsoTriState.msoTrue);
            PowerPoint.Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            // add a commandbar popup
            Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)_powerApplication.CommandBars["Menu Bar"].Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
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
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            Clipboard.SetDataObject(this.Icon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            #endregion

            #region Create a new toolbar

            // add a new toolbar
            commandBar = _powerApplication.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, false, true);
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
            Clipboard.SetDataObject(this.Icon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            #endregion

            #region Create a new ContextMenu

            // add a commandbar popup
            commandBarPopup = (Office.CommandBarPopup)_powerApplication.CommandBars["Frames"].Controls.Add(
                                                            MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarPopup.Caption = "commandBarPopup";

            // add a button to the popup
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            commandBarBtn.FaceId = 9;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            #endregion

            // make visible & set buttons
            _powerApplication.Visible = MsoTriState.msoTrue;
            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _powerApplication.Quit();
            _powerApplication.Dispose();

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
