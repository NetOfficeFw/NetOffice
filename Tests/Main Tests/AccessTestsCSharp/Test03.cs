using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;

namespace AccessTestsCSharp
{
    public class Test03 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test03"; }
        }

        public string Description
        {
            get { return "Create a table and perform CompactDatabase."; }
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
                application = new Access.Application();

                // create database name 
                string fileExtension = GetDefaultExtension(application);
                string documentFile = string.Format("{0}\\Test3{1}", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), fileExtension);

                // delete old database if exists
                if (System.IO.File.Exists(documentFile))
                    System.IO.File.Delete(documentFile);

                // create database 
                DAO.Database newDatabase = application.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);
                application.DBEngine.Workspaces[0].Close();

                // setup database connection                        'Provider=Microsoft.Jet.OLEDB.4.0;Data Source= < access2007
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
                string newDocumentFile = string.Format("{0}\\CompactDatabase{1}", Environment.CurrentDirectory, fileExtension);
                if (File.Exists(newDocumentFile))
                    File.Delete(newDocumentFile);

                application.DBEngine.CompactDatabase(documentFile, newDocumentFile);
            

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
