using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using System.Globalization;
using ExampleBase;

using NetOffice;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using NetOffice.AccessApi.Constants;
using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;

namespace AccessExamplesCS4
{
    /// <summary>
    /// Example 1 - Create new Database
    /// </summary>
    internal class Example01 : IExample 
    {
        #region IExample

        public void RunExample()
        {
            // start access 
            Access.Application accessApplication = new Access.Application();
             
            // create database name 
            string fileExtension = GetDefaultExtension(accessApplication);
            string documentFile = string.Format("{0}\\Example01{1}", HostApplication.RootDirectory, fileExtension);

            // delete old database if exists
            if (System.IO.File.Exists(documentFile))
                System.IO.File.Delete(documentFile);

            // create database 
            DAO.Database newDatabase = accessApplication.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);

            // close access and dispose reference
            accessApplication.Quit(AcQuitOption.acQuitSaveAll);
            accessApplication.Dispose();

            // show dialog for the user(you!)
            HostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return HostApplication.LCID == 1033 ? "Example01" : "Beispiel01"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Create new Database" : "Eine neue Datenbank erstellen"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

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

            double Version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (Version >= 12.00)
                return ".accdb";
            else
                return ".mdb";
        }

        #endregion
    }
}
