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

namespace Example03
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

            // Create a new MailItem.
            Outlook.MailItem mailItem = outlookApplication.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;

            // prepare item and send
            mailItem.Recipients.Add(textBoxReciever.Text);
            mailItem.Subject = textBoxSubject.Text;
            mailItem.Body = textBoxBody.Text;
            mailItem.Send();

            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();

            MessageBox.Show(this, "Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
