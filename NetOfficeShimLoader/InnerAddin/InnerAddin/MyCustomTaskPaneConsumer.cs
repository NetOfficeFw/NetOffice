using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Tools;
using NetOffice.Tools;
using System.ComponentModel;

namespace InnerAddin
{
    [CustomPane(typeof(InnerAddinPane), "InnerAddin Pane", true)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class MyCustomTaskPaneConsumer : NetOffice.OfficeApi.Tools.CustomTaskPaneConsumer
    {
        public MyCustomTaskPaneConsumer(Connect parent) : base(parent)
        {
        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Trace.WriteLine("MyCustomTaskPaneConsumer OnError " + methodKind.ToString() + " " + exception.ToString());
            MessageBox.Show("On Error " + methodKind.ToString() + " " + exception.Message);
        }

        protected override void OnCTPFactoryAvailable(Office.ICTPFactory ctpFactoryInst)
        {
            try
            {
                Trace.WriteLine("OnCTPFactoryAvailable");
                base.OnCTPFactoryAvailable(ctpFactoryInst);

                if (Parent.TaskPanes.Count > 0)
                {
                    bool paneExists = null != Parent.TaskPanes[0].Pane;
                    bool paneCreated = Parent.TaskPanes[0].Create();
                    Trace.WriteLine("Parent.TaskPanes.Count > 0");
                    Trace.WriteLine("paneExists: " + paneExists.ToString());
                    Trace.WriteLine("paneCreated: " + paneCreated.ToString());

                    //MessageBox.Show(String.Format("Pane Exists {0} Pane Created {1}", paneExists, paneCreated));
                }
                else
                {
                    Trace.WriteLine("Parent.TaskPanes.Count == 0");
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("OnCTPFactoryAvailable Exception");
                Trace.WriteLine("OnCTPFactoryAvailable Exception " + exception.ToString());
                throw;
            }
        }
    }
}
