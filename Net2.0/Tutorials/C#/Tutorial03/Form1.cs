using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

namespace Tutorial03
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

            // create new Workbook & attach close event trigger
            Excel.Workbook book = application.Workbooks.Add();
            book.BeforeCloseEvent += new Excel.Workbook_BeforeCloseEventHandler(book_BeforeCloseEvent);

            // we dispose the instance. the parameter false signals to api dont release the event listener
            // set parameter to true and the event listener will stopped and you dont get events for the instance
            // the DisposeChildInstances() method has the same method overload
             book.Close();
	book.Dispose(false);

            application.Quit();
            application.Dispose();
            /*
            * the application object is ouer root object
            * dispose them release himself and any childs of application, in this case workbooks and workbook
            * the excel instance are now removed from process list
            */

            MessageBox.Show(this, "Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void book_BeforeCloseEvent(ref bool Cancel)
        {
            
        }

        private void linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Control ctrl = sender as Control;
            if (null != ctrl)
                System.Diagnostics.Process.Start(ctrl.Tag as string);
        }
    }
}
