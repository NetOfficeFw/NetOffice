using System;
using System.Reflection; 
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;

using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using NetOffice.AccessApi.Constants;

using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;


using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;

namespace Example05
{
    public partial class Form1 : Form
    {
        Access.Application _accessApplication;

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

            Office.CommandBarButton commandBarBtn;

            // start access
            _accessApplication = new Access.Application();

            // create database name 
            string fileExtension = GetDefaultExtension(_accessApplication);
            string documentFile = string.Format("{0}\\Example05{1}", Environment.CurrentDirectory, fileExtension);


            // delete old database if exists
            if (System.IO.File.Exists(documentFile))
                System.IO.File.Delete(documentFile);

            // create database 
            DAO.Database newDatabase = _accessApplication.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);
        

            // add a commandbar popup
            Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)_accessApplication.CommandBars["Menu Bar"].Controls.Add(
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

            // make visible
            _accessApplication.Visible = true;
            button1.Enabled = false ;
            button2.Enabled = true; 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _accessApplication.Quit(AcQuitOption.acQuitSaveNone);
            _accessApplication.Dispose();

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

        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".mdb" or ".mdbx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Access.Application application)
        {
            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".mdbx";
            else
                return ".mdb";
        }

        #endregion
    }
}
