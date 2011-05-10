using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

namespace Tutorial07
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
            
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;
            application.Workbooks.Add();

            Excel.Worksheet sheet = (Excel.Worksheet)application.Workbooks[1].Worksheets[1];
            Excel.Range sampleRange = sheet.Cells[1, 1];

            // we set the COMVariant ColorIndex from Font of ouer sample range with the invoker class
            Invoker.PropertySet(sampleRange.Font, "ColorIndex", 1);

            // creates a native unmanaged ComProxy with the invoker
            object comProxy = Invoker.MethodReturn(application, "Workbooks");
            Marshal.ReleaseComObject(comProxy);

            application.Quit();
            application.Dispose();
        }
    }
}
