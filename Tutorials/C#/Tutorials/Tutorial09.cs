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
    public class Tutorial09 : ITutorial
    {
        #region ITutorial

        public void Run()
        {
            // In some situations you want use NetOffice with a already running application.
            // this examples show you how its possible.

            // GetActiveInstance take the first instance in memory
            Excel.Application excelApplication = Excel.Application.GetActiveInstance();

            // another method is GetActiveInstances:
            // 
            // GetActiveInstances takes all instances in memory. dont forget to dispose the instances.
            //            
            // Excel.Application[] excelApplications = Excel.Application.GetActiveInstances();

            excelApplication.Quit();
            excelApplication.Dispose();

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
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial09_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial09_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial09"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Create a NetOffice Excel Application Object with given COM Proxy" : "Ein NetOffice Excel Objekt Application basierend auf einem COM Proxy erstellen"; }
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
