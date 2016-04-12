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
using NetOffice.AccessApi.Tools.Utils;

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

            // create a utils instance, not need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(accessApplication);

            // create database file name 
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example01", Access.Tools.DocumentFormat.Normal);

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
    }
}
