using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Text;
using System.Data;
using System.Data.OleDb;
using ExampleBase;

using NetOffice;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;

namespace AccessExamplesCS4
{
    /// <summary>
    /// Example 5 - Customize UI
    /// </summary>
    internal partial class Example05 : UserControl, IExample
    {
        #region Fields/Delegates

        Access.Application _accessApplication;
        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        #endregion

        #region Ctor

        public Example05()
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
            get { return HostApplication.LCID == 1033 ? "Example05" : "Beispiel05"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Customize UI" : "Erweitern der klassischen Oberfläche"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region Methods

        private void UpdateTextbox(string Message)
        {
            textBoxEvents.AppendText(Message + "\r\n");
        }

        #endregion

        #region Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // start access
            _accessApplication = new Access.Application();
            Access.Tools.CommonUtils utils = new Access.Tools.CommonUtils(_accessApplication);
            Office.CommandBarButton commandBarBtn = null;

            // add a commandbar popup
            Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)_accessApplication.CommandBars["Menu Bar"].Controls.Add(MsoControlType.msoControlPopup, null, null, null, true);
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
            commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, null, null, null, true);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            Clipboard.SetDataObject(HostApplication.DisplayIcon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

            #endregion

            // make visible
            _accessApplication.Visible = true;
            buttonStartExample.Enabled = false;
            buttonQuitExample.Enabled = true;
        }

        private void buttonQuitExample_Click(object sender, EventArgs e)
        {
            _accessApplication.Quit(AcQuitOption.acQuitSaveNone);
            _accessApplication.Dispose();

            buttonStartExample.Enabled = true;
            buttonQuitExample.Enabled = false;
        }

        private void commandBarBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Click called." });
            Ctrl.Dispose();
        }

        #endregion
    }
}
