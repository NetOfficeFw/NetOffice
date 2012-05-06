using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public partial class Tutorial07 : ITutorial
    {
        IHost _hostApplication;

        #region ITutorial Member

        public void Run()
        {
            // create new instance
            Excel.Application application = new Excel.Application();

            // any reference type in NetOffice implements the EntityIsAvailable method.
            // you check here your property or method is available.

            // we check the support for 2 properties  at runtime
            bool enableLivePreviewSupport = application.EntityIsAvailable("EnableLivePreview");
            bool openDatabaseSupport = application.Workbooks.EntityIsAvailable("OpenDatabase");

            string result = "Excel Runtime Check: " + Environment.NewLine;
            result += "Support EnableLivePreview: " + enableLivePreviewSupport.ToString() + Environment.NewLine;
            result += "Support OpenDatabase:      " + openDatabaseSupport.ToString() + Environment.NewLine;
            
            // quit and dispose
            application.Quit();
            application.Dispose();

            _hostApplication.ShowMessage(result);
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public void Disconnect()
        {

        }

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return _hostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial07_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial07_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial07"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Versionindependent Development" : "Versionsunabhängige Entwicklung"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
