using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

namespace Tutorial08
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // Initialize NetOffices
            LateBindingApi.Core.Factory.Initialize();

            // create new instance
            Excel.Application application = new Excel.Application();
            
            // check for support at runtime
            bool enableLivePreviewSupport = application.EntityIsAvailable("EnableLivePreview");
            bool openDatabaseSupport = application.Workbooks.EntityIsAvailable("OpenDatabase");

            string result = "Excel Runtime Check: " + Environment.NewLine;
            result += "Support EnableLivePreview: " + enableLivePreviewSupport.ToString() + Environment.NewLine;
            result += "Support OpenDatabase:      " + openDatabaseSupport.ToString() + Environment.NewLine;

            richTextBoxResult.Text = result;

            // quit and dispose
            application.Quit();
            application.Dispose();
        }
    }
}
