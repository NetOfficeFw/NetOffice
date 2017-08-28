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

namespace ExcelTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            try
            {
                Type excelType = System.Type.GetTypeFromProgID("Excel.Application", true);
                object interopProxy = Activator.CreateInstance(excelType);

                dynamic application = new COMDynamicObject(interopProxy);
                application.Visible = true;
                application.Workbooks.Add();
                System.Threading.Thread.Sleep(10000);
                application.Quit();
                application.Dispose();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }

            //using (Excel.Application application = new NetOffice.ExcelApi.Application())
            //{
            //    application.Visible = true;
            //    var books = application.Workbooks;
            //    Excel.Workbook book = books.Add();
            //    Excel.Workbook theBook = books[1];
            //    application.Quit();
            //}
        }
    }
}
