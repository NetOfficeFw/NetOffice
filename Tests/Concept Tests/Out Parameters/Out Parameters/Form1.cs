using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Word = NetOffice.WordApi;

namespace Out_Parameters
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            LateBindingApi.Core.Factory.Initialize();

            Word.Application application = new Word.Application();
            application.Visible = true;
            Word.Document document = application.Documents.Add();
            application.Selection.TypeText("Hello World");

            int left=0;
            int top=0;
            int width=0;
            int height=0;

            application.ActiveWindow.GetPoint(out left, out top, out width, out height, application.Selection.Range);

            MessageBox.Show(string.Format("GetPoint returns Left:{0} Top:{1} Width:{2} Height:{3}", left, top, width, height));

            application.Quit();
            application.Dispose();

        }
    }
}
