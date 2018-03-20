using System;
using System.Windows.Forms;
using TutorialsBase;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial06 : ITutorial
    {
        public void Run()
        {
            // this examples shows how to use variant types(object in NetOffice) at runtime
            // the reason for the most variant definitions in office is a more flexible value set.(95%)
            // here is some code to demonstrate this

            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            // create new Workbook and a new named style
            Excel.Workbook book = application.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];
            Excel.Range range = sheet.Cells[1, 1];
            Excel.Style myStyle = book.Styles.Add("myUniqueStyle");

            // Range.Style is defined as Variant in Excel and as object in NetOffice
            // You got always an Excel.Style instance if you ask for
            Excel.Style style = (Excel.Style)range.Style;

            // and here comes the magic. both sets are valid because the variants was very flexible in the setter
            range.Style = "myUniqueStyle";
            range.Style = myStyle;

            // Name, Bold, Size are string, bool and double but defined as Variant
            style.Font.Name = "Arial";
            style.Font.Bold = true; // you can also set "true" and it works. variants makes it possible
            style.Font.Size = 14;

            // quit & dipose
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
            get { return Program.DocumentationBase + "Tutorial06_EN_CS.html"; }
        }

        public string Caption
        {
            get { return "Tutorial06"; }
        }

        public string Description
        {
            get { return "Understanding Variants"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}