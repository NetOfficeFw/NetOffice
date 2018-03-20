using System;
using System.Windows.Forms;
using TutorialsBase;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using OfficeContribution = NetOffice.OfficeApi.Tools.Contribution;
using NetOffice.ExcelApi.Tools.Contribution;

namespace TutorialsCS4
{
    public class Tutorial17 : ITutorial
    {
        public void Run()
        {
            // Any MS-Office application in NetOffice has a custom contribution provider for common tasks
            // Moreover its available as instance property in NetOffice.Tools.COMAddin
            // If you have suggestions for the contribution please feel free to contact the project
            // This tutorial shows only few features in MS-Excel

            // start excel and disable alerts
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;


            // Create an instance of excel utils
            CommonUtils utils = new CommonUtils(application, typeof(Tutorial17).Assembly);

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
            var userResult =
                utils.Dialog.ShowMessageBox(
                    "Hello World from NetOffice tutorial", "NO tutorial",
                    OfficeContribution.DialogUtils.Buttons.YesNo,
                    OfficeContribution.DialogUtils.Result.No);


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

        public string Uri
        {
            get { return Program.DocumentationBase + "Tutorial17_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial17"; }
        }

        public string Description
        {
            get { return "NetOffice Contribution"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}