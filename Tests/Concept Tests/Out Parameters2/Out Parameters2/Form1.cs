using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using NetOffice;
using Visio = NetOffice.VisioApi;
using NetOffice.VisioApi.Enums;

namespace Out_Parameters2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            Visio.Application application = new Visio.ApplicationClass();
            application.Visible = true;
            var doc = application.Documents.Add("");
            Visio.IVPage page = application.ActivePage;
            var shape = page.DrawRectangle(0, 0, 2, 3);
            shape.Text = "With Microsoft.Office.Interop.Visio";
            doc.Saved = true;

            var SID_SRCStream = new short[4];
            SID_SRCStream[0] = (short)shape.ID16;
            SID_SRCStream[1] = (short)VisSectionIndices.visSectionObject;
            SID_SRCStream[2] = (short)VisRowIndices.visRowFill;
            SID_SRCStream[3] = (short)VisCellIndices.visFillForegnd;

            try
            {
                object[] a = null;// new Array[4];
                page.GetFormulas(SID_SRCStream, out a);
                // page.GetFormulas(SID_SRCStream, out a);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
            }
            try
            {
                application.Quit();
                application.Dispose();
            }
            catch
            {
                // may closed by user
            }
        }
    }
}
