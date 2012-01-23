using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums; 

namespace Example04
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // Create an Outlook Application object. 
            Outlook.Application outlookApplication = new Outlook.Application();

            // SendAndReceive is supported from Outlook 2007 or higher
            // we check at runtime the feature is available
            if(outlookApplication.Session.EntityIsAvailable("SendAndReceive"))
            {
                // one simple call
                outlookApplication.Session.SendAndReceive(false);
            }
            else
            {
                MessageBox.Show(this, "This version of MS-Outlook doesnt support SendAndReceive.", "Example04", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }

            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();

            MessageBox.Show(this, "Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
