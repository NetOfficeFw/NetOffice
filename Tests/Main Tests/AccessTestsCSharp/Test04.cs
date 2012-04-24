using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;

namespace AccessTestsCSharp
{
    public class Test04 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test04"; }
        }

        public string Description
        {
            get { return "Create custom UI."; }
        }

        public string OfficeProduct
        {
            get { return "Access"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Access.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                Bitmap iconBitmap = new Bitmap(System.Reflection.Assembly.GetAssembly(this.GetType()).GetManifestResourceStream("AccessTestsCSharp.Test04.bmp"));
                application = new Access.Application();

                Office.CommandBarButton commandBarBtn;

                // create database name 
                string fileExtension = GetDefaultExtension(application);
                string documentFile = string.Format("{0}\\Test4{1}", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), fileExtension);

                // delete old database if exists
                if (System.IO.File.Exists(documentFile))
                    System.IO.File.Delete(documentFile);

                // create database 
                DAO.Database newDatabase = application.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);

                // add a commandbar popup
                Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)application.CommandBars["Menu Bar"].Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarPopup.Caption = "commandBarPopup";
                  
                #region CommandBarButton

                // add a button to the popup
                commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
                commandBarBtn.Caption = "commandBarButton";
                Clipboard.SetDataObject(iconBitmap);
                commandBarBtn.PasteFace();
                commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

                #endregion

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    // close access and dispose reference
                    application.Quit(AcQuitOption.acQuitSaveNone);
                    application.Dispose();
                }
            }
        }

        void commandBarBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Ctrl.Dispose();
        }

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
