using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using Word = NetOffice.WordApi;
using Outlook = NetOffice.OutlookApi;
using PowerPoint = NetOffice.PowerPointApi;
using Access = NetOffice.AccessApi;
using Project = NetOffice.MSProjectApi;
using Visio = NetOffice.VisioApi;
using Office = NetOffice.OfficeApi;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeUI
{
    /// <summary>
    /// Office product application wrapper to create and dispose/get commandbars
    /// </summary>
    internal class ApplicationWrapper
    {
        #region Fields

        private string _officeApp;
        private Excel.Application _excelApplication;
        private Word.Application _wordApplication;
        private Outlook.Application _outlookApplication;
        private PowerPoint.Application _powerpointApplication;
        private Access.Application _accessApplication;
        private Project.Application _projectApplication;
        private Visio.Application _visioApplication;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="officeApp">name of the office application to create</param>
        internal ApplicationWrapper(string officeApp)
        {
            _officeApp = officeApp;
            CreateOfficeApplication();
        }

        #endregion

        #region Properties

        /// <summary>
        /// CommandBars from the application user interface
        /// </summary>
        public Office.CommandBars CommandBars
        {
            get
            {
                switch (_officeApp)
                {
                    case "Excel":
                        return _excelApplication.CommandBars;
                    case "Word":
                        return _wordApplication.CommandBars;
                    case "Outlook":
                         Outlook._NameSpace outlookNS = _outlookApplication.GetNamespace("MAPI");
                         Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OutlookApi.Enums.OlDefaultFolders.olFolderInbox);
                         inboxFolder.Display();
                        return (_outlookApplication.ActiveWindow() as OutlookApi.Explorer) .CommandBars;
                    case "Power Point":
                        return _powerpointApplication.CommandBars;
                    case "Access":
                        return _accessApplication.CommandBars;
                    case "Project":
                        return _projectApplication.CommandBars;
                    case "Visio":
                        return _visioApplication.CommandBars as Office.CommandBars;
                    default:
                        throw new ArgumentOutOfRangeException("officeApp");
                }
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Quit application
        /// </summary>
        public void Quit()
        {
            switch (_officeApp)
            {
                case "Excel":
                    _excelApplication.Quit();
                    break;
                case "Word":
                    _wordApplication.Quit();
                    break;
                case "Outlook":
                    _outlookApplication.Quit();
                    break;
                case "Power Point":
                    _powerpointApplication.Quit();
                    break;
                case "Access":
                    _accessApplication.Quit();
                    break;
                case "Project":
                    _projectApplication.Quit();
                    break;
                case "Visio":
                    _visioApplication.Quit();
                    break;
                default:
                    throw new ArgumentOutOfRangeException("officeApp");
            }
        }

        /// <summary>
        /// Dispose application
        /// </summary>
        public void Dispose()
        {
            switch (_officeApp)
            {
                case "Excel":
                    _excelApplication.Dispose();
                    break;
                case "Word":
                    _wordApplication.Dispose();
                    break;
                case "Outlook":
                    _outlookApplication.Dispose();
                    break;
                case "Power Point":
                    _powerpointApplication.Dispose();
                    break;
                case "Access":
                    _accessApplication.Dispose();
                    break;
                case "Project":
                    _accessApplication.Dispose();
                    break;
                case "Visio":
                    _accessApplication.Dispose();
                    break;
                default:
                    throw new ArgumentOutOfRangeException("officeApp");
            }
        }

        /// <summary>
        /// Create application
        /// </summary>
        private void CreateOfficeApplication()
        {
            switch (_officeApp)
            {
                case "Excel":
                    _excelApplication = new Excel.ApplicationClass("Excel.Application");
                    break;
                case "Word":
                    _wordApplication = new Word.ApplicationClass("Word.Application");
                    break;
                case "Outlook":
                    _outlookApplication = new Outlook.ApplicationClass("Outlook.Application");
                    break;
                case "Power Point":
                    _powerpointApplication = new PowerPoint.ApplicationClass("PowerPoint.Application");
                    break;
                case "Access":
                    _accessApplication = new Access.ApplicationClass("Access.Application");
                    break;
                case "Project":
                    _projectApplication = new Project.ApplicationClass("MSProject.Application");
                    break;
                case "Visio":
                    _visioApplication = new Visio.ApplicationClass("Visio.Application");
                    break;
                default:
                    throw new ArgumentOutOfRangeException("officeApp");
            }
        }

        #endregion
    }
}
