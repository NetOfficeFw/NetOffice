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
using Office = NetOffice.OfficeApi;

namespace TutorialsCS4
{
    public class Tutorial05 : ITutorial
    { 
        #region ITutorial Member

        public void Run()
        {
            // this is a simple demonstration how i can convert unkown types at runtime

            // start application
            Excel.Application application = new Excel.Application();
            application.Visible = false;
            application.DisplayAlerts = false;

            foreach (Office.COMAddIn item in application.COMAddIns)
            {
                // the application property is an unkown COM object
                // we need a cast at runtime
                Excel.Application hostApp = item.Application as Excel.Application;
                
                // do some sample stuff
                string hostAppName = hostApp.Name;
                bool hostAppVisible = hostApp.Visible;
            }
 
            // quit and dispose excel
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
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial05_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial05_DE_CS"; }

        }

        public string Caption
        {
            get { return "Tutorial05"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Understanding unkown Types" : "Richtiges verwenden von unbekannten COM Objekten"; }
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
