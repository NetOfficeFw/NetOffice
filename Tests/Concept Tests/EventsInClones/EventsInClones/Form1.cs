using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace EventsInClones
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Test();
        }

        /// <summary>
        ///
        /// </summary>
        private void Test()
        {
            try
            {
                using (Excel.Application application = new Excel.ApplicationClass())
                {
                    application.DisplayAlerts = false;
                    application.NewWorkbookEvent += Application1_NewWorkbookEvent;
                    using (Excel.Application application2 = application.DeepCopy())
                    {
                        application2.NewWorkbookEvent += Application2_NewWorkbookEvent;
                        application.Workbooks.Add();
                    }

                    application.Quit();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
            }
        }

        private void Application1_NewWorkbookEvent(Excel.Workbook wb)
        {
            MessageBox.Show("Application1_NewWorkbookEvent");
        }

        private void Application2_NewWorkbookEvent(Excel.Workbook wb)
        {
            MessageBox.Show("Application2_NewWorkbookEvent");
        }
    }
}
