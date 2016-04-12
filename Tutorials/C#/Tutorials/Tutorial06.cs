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
    public class Tutorial06 : ITutorial 
    {
        #region ITutorial

        public void Run()
        {
            // this examples shows how i can use variant types(object in NetOffice) at runtime
            // the reason for the most variant definitions in office is a more flexible value set.(95%)
            // here is the code to demonstrate this

            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            // create new Workbook and a new named style
            Excel.Workbook book = application.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];
            Excel.Range range = sheet.Cells[1, 1];
            Excel.Style myStyle = book.Styles.Add("myUniqueStyle");
 
            // Range.Style is defined as Variant in Excel and represents as object in NetOffice
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

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial06_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial06_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial06"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Understanding Variant" : "Verstehen und verwenden von Variant Typen"; }
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
