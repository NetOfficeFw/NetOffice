using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial10 : ITutorial
    {
        public void Run()
        {
            // Enable and trigger trace alert
            NetOffice.Settings.Default.PerformanceTrace.Enabled = true;
            NetOffice.Settings.Default.PerformanceTrace.Alert += delegate(NetOffice.PerformanceTrace sender, NetOffice.PerformanceTrace.PerformanceAlertEventArgs args)
            {
                Console.WriteLine("{0} {1}:{2} in {3} Milliseconds ({4} Ticks)", args.CallType, args.EntityName, args.MethodName, args.TimeElapsedMS, args.Ticks);
            };

            // Criteria 1
            // Enable performance trace in excel generaly. set interval limit to 100ms to see all actions there need >= 100 milliseconds
            NetOffice.Settings.Default.PerformanceTrace["NetOffice.ExcelApi"].Enabled = true;
            NetOffice.Settings.Default.PerformanceTrace["NetOffice.ExcelApi"].IntervalMS = 100;

            // Criteria 2
            // Enable additional performance trace for all members of WorkSheet in excel. set interval limit to 20ms to see all actions there need >=20 milliseconds
            NetOffice.Settings.Default.PerformanceTrace["NetOffice.ExcelApi", "Worksheet"].Enabled = true;
            NetOffice.Settings.Default.PerformanceTrace["NetOffice.ExcelApi", "Worksheet"].IntervalMS = 20;

            // Criteria 3
            // Enable additional performance trace for WorkSheet Range property in excel. set interval limit to 0ms to see all calls anywhere
            NetOffice.Settings.Default.PerformanceTrace["NetOffice.ExcelApi", "Worksheet", "Range"].Enabled = true;
            NetOffice.Settings.Default.PerformanceTrace["NetOffice.ExcelApi", "Worksheet", "Range"].IntervalMS = 0;

            // do some stuff
            Excel.Application application = new NetOffice.ExcelApi.Application();
            application.DisplayAlerts = false;
            Excel.Workbook book = application.Workbooks.Add();
            Excel.Worksheet sheet = book.Sheets.Add() as Excel.Worksheet;
            for (int i = 1; i <= 5; i++)
            {
                Excel.Range range = sheet.Range("A" + i.ToString());
                range.Value = "Test123";
                range[1, 1].Value = "Test234";
            }
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
            get { return Program.DocumentationBase + "Tutorial10_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial10"; }
        }

        public string Description
        {
            get { return "Measure Performance"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}