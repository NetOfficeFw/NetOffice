using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;

using LateBindingApi.Core;

using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using NetOffice.AccessApi.Constants;

using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;

namespace Example03
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // start access 
            Access.Application accessApplication = new Access.Application();

            // create database name 
            string fileExtension = GetDefaultExtension(accessApplication);
            string documentFile = string.Format("{0}\\Example03{1}", Application.StartupPath, fileExtension);


            // delete old database if exists
            if (System.IO.File.Exists(documentFile))
                System.IO.File.Delete(documentFile);

            // create database 
            DAO.Database newDatabase = accessApplication.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);
            accessApplication.DBEngine.Workspaces[0].Close();

            // setup database connection
            OleDbConnection oleConnection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + documentFile);
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

            accessApplication.DBEngine.CompactDatabase(documentFile, newDocumentFile);
            
            // close access and dispose reference
            accessApplication.Quit(AcQuitOption.acQuitSaveAll);
            accessApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Database saved.", newDocumentFile);
            fDialog.ShowDialog(this);
        }

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
