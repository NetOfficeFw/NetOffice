using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = LateBindingApi.ExcelApi;

namespace Tutorial05
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

            // ActiveSheet is defined as unkown Proxy in Excel Type Library, it can have multiple times at runtime
            // In VBA oder PIA its converted to object, in LateBindingApi its represents as COMObject
            // All LateBindingApi Classes inherited COMObject
            COMObject sheet = application.ActiveSheet;
            if (sheet is Excel.Worksheet)
            {
                Excel.Worksheet activeSheet = (Excel.Worksheet)sheet;
            }

            // 3 basic properties of COMObject
            object proxy          = sheet.UnderlyingObject;
            string proxyClassName = sheet.UnderlyingTypeName;
            bool   isDisposed     = sheet.IsDisposed;
 
            application.Quit();
            application.Dispose();

            MessageBox.Show(this, "Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
