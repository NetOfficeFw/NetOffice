using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using ExampleBase;

using NetOffice;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;

namespace AccessExamplesCS4
{
    class Example03 : IExample
    {
        IHost _hostApplication;

        #region IExample Member

        public void RunExample()
        {
            // start access 
            Access.Application accessApplication = new Access.Application();

            // create database name 
            string fileExtension = GetDefaultExtension(accessApplication);
            string documentFile = string.Format("{0}\\Example03{1}", _hostApplication, fileExtension);


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

            // now we do CompactDatabase            

            string newDocumentFile = string.Format("{0}\\CompactDatabase{1}", _hostApplication, fileExtension);
            if (File.Exists(newDocumentFile))
                File.Delete(newDocumentFile);

            accessApplication.DBEngine.CompactDatabase(documentFile, newDocumentFile);

            // close access and dispose reference
            accessApplication.Quit(AcQuitOption.acQuitSaveAll);
            accessApplication.Dispose();

            // show dialog for the user(you!)
            _hostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example03" : "Beispiel03"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Use CompactDatabase" : "CompactDatabase ausführen"; }
        }

        public UserControl Panel
        {
            get { return null; }
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

            double Version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (Version >= 12.00)
                return ".accdb";
            else
                return ".mdb";
        }

        #endregion
    }
}
