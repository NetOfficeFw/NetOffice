using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

namespace Tutorial02
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

            Excel.Workbook book   = application.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets.Add();

            /*
             * we have 5 created proxies now in proxy table as follows
             * 
             * Application
             *   + Workbooks
             *     + Workbook  
             *        + Worksheets  
             *            + Worksheet  
            */


            // we dispose the child instances of book
            book.DisposeChildInstances();

            /*
            * we have 3 created proxies now, the childs from book are disposed
            * 
            * Application
            *   + Workbooks
            *     + Workbook  
            */

            application.Quit();
            application.Dispose();
            /*
            * the Dispose() call for application release the instance and created childs Workbooks and Workbook
            */

            MessageBox.Show(this, "Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
