using System;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using ExampleBase;
using NetOffice;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;
using NetOffice.AccessApi.Tools.Contribution;

namespace AccessExamplesCS4
{
    /// <summary>
    /// Example 3 - Use CompactDatabase
    /// </summary>
    internal class Example03 : IExample
    {
        public void RunExample()
        {
            // start access 
            Access.Application accessApplication = new Access.Application();

            // create a utils instance, no need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(accessApplication);

            // create database file name 
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example03", DocumentFormat.Normal);

            // delete old database if exists
            if (System.IO.File.Exists(documentFile))
                System.IO.File.Delete(documentFile);

            // create database 
            DAO.Database newDatabase = accessApplication.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);
            accessApplication.DBEngine.Workspaces[0].Close();
 
            // setup database connection                         'Provider=Microsoft.Jet.OLEDB.4.0;Data Source= < access2007
            OleDbConnection oleConnection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False;Data Source=" + documentFile);
            oleConnection.Open();

            // create table
            OleDbCommand oleCreateCommand = new OleDbCommand("CREATE TABLE NetOfficeTable(Column1 Text, Column2 Text)", oleConnection);
            oleCreateCommand.ExecuteReader().Close();

            // write some data with plain sql & close
            for (int i = 0; i < 20000; i++)
            {
                string insertCommand = string.Format("INSERT INTO NetOfficeTable(Column1, Column2) VALUES(\"{0}\", \"{1}\")", i, DateTime.Now.ToShortTimeString());
                OleDbCommand oleInsertCommand = new OleDbCommand(insertCommand, oleConnection);
                oleInsertCommand.ExecuteReader().Close();
            }
            oleConnection.Close();

            // delete old file if exists
            string newDocumentFile = string.Format("{0}\\CompactDatabase{1}", HostApplication, utils.File.FileExtension(DocumentFormat.Normal));
            if (File.Exists(newDocumentFile))
                File.Delete(newDocumentFile);

            // now we do CompactDatabase            
            accessApplication.DBEngine.CompactDatabase(documentFile, newDocumentFile);

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
            get { return "Example03"; }
        }

        public string Description
        {
            get { return "Use CompactDatabase"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
