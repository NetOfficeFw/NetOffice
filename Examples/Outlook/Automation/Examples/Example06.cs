using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookExamplesCS4
{
    /// <summary>
    /// Example 6 - Using events
    /// </summary>
    internal partial class Example06 : UserControl, IExample
    {
        #region Fields/Delegates

        private delegate void UpdateEventTextDelegate(string Message);
        private UpdateEventTextDelegate _updateDelegate;

        #endregion

        #region Ctor

        public Example06()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
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
            get { return "Example06"; }
        }

        public string Description
        {
            get { return "Events"; }
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

        #region Methods

        private void UpdateTextbox(string message)
        {
            textBoxEvents.AppendText(message + "\r\n");
        }

        #endregion

        #region Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // start outlook by trying to access running application first
            Outlook.Application outlookApplication = new Outlook.Application(true);

            // create MailItem and register close event
            Outlook.MailItem mailItem = outlookApplication.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;
            mailItem.CloseEvent += new NetOffice.OutlookApi.MailItem_CloseEventHandler(mailItem_CloseEvent);

            // BodyFormat is not available in Outlook 2000, we check at runtime the property is available
            if (mailItem.EntityIsAvailable("BodyFormat"))
                mailItem.BodyFormat = OlBodyFormat.olFormatPlain;
            mailItem.Body = "ExampleBody";
            mailItem.Subject = "ExampleSubject";
            mailItem.Display();
            mailItem.Close(OlInspectorClose.olDiscard);

            // close outlook and dispose
            if (!outlookApplication.FromProxyService)
                outlookApplication.Quit();
            outlookApplication.Dispose();
        }

        private void mailItem_CloseEvent(ref bool Cancel)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event Close called." });
        }

        #endregion
    }
}
