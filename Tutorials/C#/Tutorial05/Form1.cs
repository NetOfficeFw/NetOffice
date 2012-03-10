using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

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
            // but its always a COM Proxy, never a scalar type like bool or int. 
            // In VBA oder PIA its converted to object, in NetOffice its represents as COMObject
            // All NetOffice classes inherited COMObject
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

        private void linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Control ctrl = sender as Control;
            if (null != ctrl)
                System.Diagnostics.Process.Start(ctrl.Tag as string);
        }
    }
}
