using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public partial class Tutorial12 : UserControl, ITutorial
    {
        public Tutorial12()
        {
            InitializeComponent();
        }

        public void Run()
        {

        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
            propertyGrid1.SelectedObject = NetOffice.Settings.Default;
        }

        public void Disconnect()
        {

        }

        public string Uri
        {
            get { return Program.DocumentationBase + "Tutorial12_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial12"; }
        }


        public string Description
        {
            get { return "NetOffice Settings"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
