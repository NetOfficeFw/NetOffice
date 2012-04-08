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

using LateBindingApi.Core;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;

namespace AccessExamplesCS4
{
    public partial class Example05 : UserControl, IExample
    {
        IHost _hostApplication;

        Access.Application _accessApplication;

        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Example05()
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
            get { return _hostApplication.LCID == 1033 ? "Example05" : "Beispiel05"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Customize UI" : "Erweitern der klassischen Oberfläche"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }

        #endregion

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            Office.CommandBarButton commandBarBtn;

            // start access
            _accessApplication = new Access.Application();

            // create database name 
            string fileExtension = GetDefaultExtension(_accessApplication);
            string documentFile = string.Format("{0}\\Example05{1}", _hostApplication, fileExtension);

            // delete old database if exists
            if (System.IO.File.Exists(documentFile))
                System.IO.File.Delete(documentFile);

            // create database 
            DAO.Database newDatabase = _accessApplication.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);

            // add a commandbar popup
            Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)_accessApplication.CommandBars["Menu Bar"].Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
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
            Clipboard.SetDataObject(_hostApplication.DisplayIcon.ToBitmap());
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

        #region Access Trigger

        void commandBarBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Click called." });
            Ctrl.Dispose();
        }

        private void UpdateTextbox(string Message)
        {
            textBoxEvents.AppendText(Message + "\r\n");
        }

        #endregion

        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".mdb" or ".accdb"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Access.Application application)
        {
            // Access 2000 doesnt have the Version property(unfortunately)
            // we check for support with the SupportEntity method, implemented by NetOffice
            if (!application.EntityIsAvailable("Version"))
                return ".mdb";

            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".accdb";
            else
                return ".mdb";
        }

        #endregion
    }
}
