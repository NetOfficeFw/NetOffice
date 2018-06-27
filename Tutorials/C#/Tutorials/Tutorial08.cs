using System;
using System.Windows.Forms;
using TutorialsBase;
using Excel = NetOffice.ExcelApi;
using NetOffice;

namespace TutorialsCS4
{
    /// <summary>
    /// The custom Excel.Workbook
    /// </summary>
    public class MyWorkbook : NetOffice.ExcelApi.Behind.Workbook
    {
        public override Excel.Sheets Sheets
        {
            get
            {
                return base.Sheets;
            }
        }

        public bool Has3Sheets
        {
            get
            {
                return Sheets.Count == 3;
            }
        }
    }

    public class Tutorial08 : ITutorial
    {
        public void Run()
        {
            // Replace Excel.Workbook with MyWorkbook
            NetOffice.Core.Default.ObjectActivator.CreateInstance += delegate(Core sender, NetOffice.CoreServices.OnCreateInstanceEventArgs args)
            {
                if (args.RequestedFrom.ContractType == typeof(Excel.Workbook))
                    args.Replace = new MyWorkbook();
            };

            Excel.Application application = new Excel.ApplicationClass();
            application.DisplayAlerts = false;

            // add and cast book to MyWorkbook
            MyWorkbook book = application.Workbooks.Add() as MyWorkbook;
            if (book.Has3Sheets)
                Console.WriteLine("Book has 3 sheets.");

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
            get { return Program.DocumentationBase + "Tutorial08_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial08"; }
        }


        public string Description
        {
            get { return "Custom Instances"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}