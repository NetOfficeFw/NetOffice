using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using ExampleBase;

using LateBindingApi.Core;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookExamplesCS4
{
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

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example03" : "Beispiel03"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Send an E- Mail" : "Eine E-Mail verschicken"; }
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
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

            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

            // create a new MailItem.
            Outlook.MailItem mailItem = outlookApplication.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;

            // prepare item and send
            mailItem.Recipients.Add(textBoxReciever.Text);
            mailItem.Subject = textBoxSubject.Text;
            mailItem.Body = textBoxBody.Text;
            mailItem.Send();

            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();

            _hostApplication.ShowFinishDialog("Done!", null);
        }

        #endregion
    }
}
