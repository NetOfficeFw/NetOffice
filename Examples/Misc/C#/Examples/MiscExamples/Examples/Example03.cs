using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using ExampleBase;

using Outlook = NetOffice.OutlookApi;

namespace MiscExamplesCS4
{
    /*
        *  in some situations you want check for a running office application instance.
        *  this example shows you how to use the Marshal.GetActiveObject method to get a running application and create a NetOffice wrapper instance.
        *  for this example we use outlook. please note the Marshal.GetActiveObject method throws a COMException if no running instance available
    */
    public partial class Example03 : UserControl, IExample
    {
        IHost _hostApplication;

        public Example03()
        {
            InitializeComponent();
        }
         
        #region IExample Member

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example03" : "Beispiel03"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "How to access a running Outlook application" : "Eine laufene Outlook Instanz automatisieren"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }

        #endregion

        #region UI Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            try
            {
                Outlook.Application application = null;

                object nativeProxy = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
                application = new Outlook.Application(null, nativeProxy);

                textBoxLog.Clear();
                textBoxLog.AppendText("we got running outlook instance\r\n");
                textBoxLog.AppendText("outlook version is " + application.Version);

                // instance was already running at start. we dispose references but not quit application
                application.Dispose();
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                _hostApplication.ShowErrorDialog(null, exception);
            }            
        }
        
        #endregion
    }
}
