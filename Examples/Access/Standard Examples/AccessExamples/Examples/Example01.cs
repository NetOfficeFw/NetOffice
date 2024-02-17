using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using NetOffice.AccessApi.Constants;
using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;
using NetOffice.AccessApi.Tools.Contribution;

namespace AccessExamplesCS4
{
    /// <summary>
    /// Example 1 - Create new Database
    /// </summary>
    internal class Example01 : IExample 
    {
        public void RunExample()
        {
            // start access 
            Access.Application accessApplication = new Access.Application();

            // create a utils instance, no need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(accessApplication);

            // create database file name 
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example01", DocumentFormat.Normal);

            // delete old database if exists
            if (System.IO.File.Exists(documentFile))
                System.IO.File.Delete(documentFile);

            // create database 
            DAO.Database newDatabase = accessApplication.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);

            // close access and dispose reference
            accessApplication.Quit(AcQuitOption.acQuitSaveAll);
            accessApplication.Dispose();

            // show end dialog
            HostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return  "Example01"; }
        }

        public string Description
        {
            get { return "Create new Database"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
