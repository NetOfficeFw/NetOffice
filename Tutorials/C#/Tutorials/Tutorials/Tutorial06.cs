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
    public partial class Tutorial06 : ITutorial 
    {
        IHost _hostApplication;

        #region ITutorial Member

        public void Run()
        {
            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            // create new Workbook
            Excel.Workbook book = application.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];
            Excel.Range range = sheet.Cells[1, 1];

            // Style is defined as Variant in Excel and represents as object in NetOffice
            // You can cast them at runtime without problems
            Excel.Style style = (Excel.Style)range.Style;

            // variant types can be a scalar type at runtime
            // another example way to use is 
            if (range.Style is string)
            {
                string myStyle = range.Style as string;
            }
            else if (range.Style is Excel.Style)
            {
                Excel.Style myStyle = (Excel.Style)range.Style;
            }

            // Name, Bold, Size are bool but defined as Variant and also converted to object
            style.Font.Name = "Arial";
            style.Font.Bold = true;
            style.Font.Size = 14;


            // Please note: the reason for the most variant definition is a more flexible value set.
            // the Style property from Range returns always a Style object
            // but if you have a new named style created with the name "myStyle" you can set range.Style = myNewStyleObject; or range.Style = "myStyle"
            // this kind of flexibility is the primary reason for Variants in Office
            // in any case, you dont lost the COM Proxy management from NetOffice for Variants. 

            // quit & dipose
            application.Quit();
            application.Dispose();

            _hostApplication.ShowFinishDialog();
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
            get { return _hostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial06_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial06_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial06"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Understanding Variant" : "Verstehen und verwenden von Variant Typen"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
