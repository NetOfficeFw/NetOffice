using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookExamplesCS4
{
    /// <summary>
    /// Example 3 - Send mail
    /// </summary>
    internal partial class Example03 : UserControl, IExample
    {
        #region Ctor

        public Example03()
        {
            InitializeComponent();
        }

        #endregion

        #region IExample Member

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public string Caption
        {
            get { return "Example03"; }
        }

        public string Description
        {
            get { return "Send an E-Mail"; }
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public UserControl Panel
        {
            get { return this; }
        }
        
        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region UI Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // start outlook by trying to access running application first
            Outlook.Application outlookApplication = new Outlook.Application(true);

            // create a new MailItem.
            Outlook.MailItem mailItem = outlookApplication.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;

            // prepare item and send
            mailItem.Recipients.Add(textBoxReciever.Text);
            mailItem.Subject = textBoxSubject.Text;
            mailItem.Body = textBoxBody.Text;
            mailItem.Send();

            // close outlook and dispose
            if (!outlookApplication.FromProxyService)
                outlookApplication.Quit();
            outlookApplication.Dispose();

            HostApplication.ShowFinishDialog("Done!", null);
        }

        #endregion
    }
}
