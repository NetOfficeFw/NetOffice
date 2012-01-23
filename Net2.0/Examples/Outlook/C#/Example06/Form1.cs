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

namespace Example06
{
    public partial class Form1 : Form
    {
        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Form1()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

            Outlook.MailItem mailItem = outlookApplication.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;
            mailItem.CloseEvent += new NetOffice.OutlookApi.MailItem_CloseEventHandler(mailItem_CloseEvent);

            // BodyFormat is not available in Outlook 2000
            // we check at runtime the property is available
            if(mailItem.EntityIsAvailable("BodyFormat"))
                mailItem.BodyFormat = OlBodyFormat.olFormatPlain;

            mailItem.Body = "Body of Example06 " + DateTime.Now.ToLongTimeString();
            mailItem.Subject = "Example06";
            mailItem.Display();
            mailItem.Close(OlInspectorClose.olDiscard);
   
            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();
        }

        void mailItem_CloseEvent(ref bool Cancel)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event Close called." });
        }

        private void UpdateTextbox(string message)
        {
            textBoxEvents.AppendText(message + "\r\n");
        }
    }
}
