using System;
using System.Collections.Generic;
using System.Text;
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
    public class Test01 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test01"; }
        }

        public string Description
        {
            get { return "Insert text."; }
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
                string documentFile = string.Format("{0}\\Test01{1}", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), fileExtension);
         
                // delete old database if exists
                if (System.IO.File.Exists(documentFile))
                    System.IO.File.Delete(documentFile);

                // create database 
                DAO.Database newDatabase = application.DBEngine.Workspaces[0].CreateDatabase(documentFile, LanguageConstants.dbLangGeneral);

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
