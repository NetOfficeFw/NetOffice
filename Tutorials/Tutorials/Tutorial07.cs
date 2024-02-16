using System;
using System.Windows.Forms;
using TutorialsBase;
using Excel = NetOffice.ExcelApi;
using NetOffice;

namespace TutorialsCS4
{
    public class Tutorial07 : ITutorial
    {
        public void Run()
        {
            HostApplication.ShowFinishDialog("Support for dynamic COM objects was removed from NetOffice.");
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
            get { return Program.DocumentationBase + "Tutorial07_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial07"; }
        }


        public string Description
        {
            get { return "Managed Dynamics (Obsolete)"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
