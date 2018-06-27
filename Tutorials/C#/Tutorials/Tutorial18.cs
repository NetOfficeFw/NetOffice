using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial18 : ITutorial
    {
        public void Run()
        {
            /*
             *  NetOffice provides features to compare 2 proxies directly on server.
             *
             *  2 proxies may different instances but pointing to the same instance on the com server(the office application)
             *
             *  This is a showstopper to demonstrate a deep comparison.
             *
             *  -------------------------------------------------------
             *  Former NetOffice versions spend operator overloads here.
             *  This is impossible in NetOffice 2.0 and above because
             *  NetOffice 2.0 use interfaces instead of classes.
             *
            */

            using (var application = COMObject.Create<Excel.Application>())
            {
                application.DisplayAlerts = false;
                Excel.Workbook book = application.Workbooks.Add();

                bool isEqual = false;

                // determine active workbook is the same as book1 on the server
                isEqual = application.ActiveWorkbook.EqualsOnServer(book);

                // another static version to do the same
                isEqual = COMObject.EqualsOnServer(application.ActiveWorkbook, book);

                application.Quit();
            }

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
            get { return Program.DocumentationBase + "Tutorial18_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial18"; }
        }

        public string Description
        {
            get { return "Compare Instances"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}