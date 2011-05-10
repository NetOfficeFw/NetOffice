using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

namespace Tutorial06
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            // create new Workbook
            Excel.Workbook book = application.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];
            Excel.Range range = sheet.Cells[1,1];

            // Style is defined as Variant in Excel Type Library and represents as object in NetOffice
            Excel.Style style = (Excel.Style)range.Style;

            // variant types can be a scalar type, another way to us is 
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
 
            // quit & dipose
            application.Quit();
            application.Dispose();

            MessageBox.Show(this, "Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
