using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;

using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using NetOffice.AccessApi.Constants;

using DAO = NetOffice.DAOApi;
using NetOffice.DAOApi.Enums;
using NetOffice.DAOApi.Constants;
 
namespace Example01
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
            string documentFile = string.Format("{0}\\Example01{1}", Environment.CurrentDirectory, fileExtension);
            
            // delete old database if exists
            if (System.IO.File.Exists(documentFile))
                System.IO.File.Delete(documentFile);
            
            // create database 
            DAO.Database newDatabase = accessApplication.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);

            // close power point and dispose reference
            accessApplication.Quit(AcQuitOption.acQuitSaveAll);
            accessApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Database saved.", (string)documentFile);
            fDialog.ShowDialog(this);
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
