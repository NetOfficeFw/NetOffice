using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.ExcelApi.Tools.Utils;

namespace TutorialsCS4
{
    public class Tutorial13 : ITutorial
    {
        #region ITutorial

        public void Run()
        {
            // Any MS-Office application in NetOffice has a custom utils provider for common tasks
            // Moreover its available as instance property in NetOffice.Tools.COMAddin
            // If you have suggestions for the utils please feel free to contact the project
            // This tutorial shows only few features in MS-Excel

            // start excel and disable alerts
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;


            // Create an instance of excel utils
            CommonUtils utils = new CommonUtils(application, typeof(Tutorial13).Assembly);


            // the file part of the utils makes it easier to deal with file extensions depedent on the current version


            // get default(xls or xlsx) , template with macros(xlt or xltm) - extension and build a valid file path
            string extensionNormal = utils.File.FileExtension(DocumentFormat.Normal);
            string extensionTemplateWithMacros = utils.File.FileExtension(DocumentFormat.TemplateMacros);
            string exampleFilePath = utils.File.Combine("C:\\MyFiles", "MyWorkbook", DocumentFormat.Normal);

            // the dialog part of the utils allows you to show default dialogs/messageboxes or you own dialogs


            // dialogs want be suppressed by default if the office application is currently in automation or not visible
            // you can also trigger the DialogShow and DialogShown event to observe dialog popups
            // we disable any suppress behavior here
            utils.Dialog.SuppressOnAutomation = false; 
            utils.Dialog.SuppressOnHide = false;


            // show a simple message box. Have a look at the last argument. Its a default result and used if the messagebox is not shown.
            // In this tutorial, excel is in automation and hidden. Remove one or both of the 2 code lines above and the message box is not shown.
            // We got the default result in this case
            DialogResult userResult = utils.Dialog.ShowMessageBox("Hello World from NetOffice tutorial", "NO tutorial", MessageBoxButtons.YesNo, DialogResult.No);


            application.Quit();
            application.Dispose();
  
            HostApplication.ShowFinishDialog();
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public void Disconnect()
        {

        }

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial13_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial13_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial13"; }
        }


        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "NetOffice Utils" : "NetOffice Utils"; }
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
