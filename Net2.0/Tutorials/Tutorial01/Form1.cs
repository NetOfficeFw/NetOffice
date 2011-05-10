using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

namespace Tutorial01
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            /*
            *  NetOffice manages COM Proxies for you to avoid any kind of memory leaks
            *  and make sure your application instance removes from process list if you want.
            */

            // Initialize Api COMObject Support 
            LateBindingApi.Core.Factory.Initialize();

            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            Excel.Workbook book = application.Workbooks.Add();
            /* 
            * now we have 2 new COM Proxies created.
            * 
            * the first proxy was created while accessing the Workbooks collection from application
            * the second proxy was created by the Add() method from Workbooks and stored now in book
            * with the application object we have 3 created proxies now. the workbooks proxy was created
            * about application and the book proxy was created about the workbooks.
            * the api holds the proxies now in a list as follows:
            * 
            * Application
            *   + Workbooks
            *     + Workbook  
            * 
            * any object in NetOffice implements the IDisposible Interface.
            * use the Dispose() Method to release an object. the method release all created child proxies too.
            */


            application.Quit();
            application.Dispose();
            /*
            * the application object is ouer root object
            * dispose them release himself and any childs of application, in this case workbooks and workbook
            * the excel instance are now removed from process list
            */

            MessageBox.Show(this, "Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
