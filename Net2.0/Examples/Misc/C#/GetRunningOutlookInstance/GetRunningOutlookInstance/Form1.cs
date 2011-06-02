using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using Outlook = NetOffice.OutlookApi; 

namespace GetRunningOutlookInstance
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            // initialize api
            LateBindingApi.Core.Factory.Initialize();  

            Outlook.Application application = null;

            object nativeProxy = RunningObjectTable.GetRunningOutlookInstanceFromROT();
            if (null != nativeProxy)
            {
                application = new NetOffice.OutlookApi.Application(null, nativeProxy);
                
                textBoxLog.Clear();
                textBoxLog.AppendText("we got running outlook instance\r\n");
                textBoxLog.AppendText("outlook version is " + application.Version );

                // instance was already running at start. we dispose references but not quit application
                application.Dispose();
            }
            else
            { 
                application = new NetOffice.OutlookApi.Application();
                
                textBoxLog.Clear();
                textBoxLog.AppendText("we create new outlook instance\r\n");
                textBoxLog.AppendText("outlook version is " + application.Version);

                // quit and dispose application
                application.Quit();
                application.Dispose();
            }
        }

    }
}
