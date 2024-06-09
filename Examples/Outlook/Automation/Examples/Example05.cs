using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookExamplesCS4
{
    /// <summary>
    /// Example 5 - Enumerate Contacts
    /// </summary>
    internal partial class Example05 : UserControl, IExample
    {
        #region Ctor

        public Example05()
        {
            InitializeComponent();
        }

        #endregion

        #region IExample

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public string Caption
        {
            get { return "Example05"; }
        }

        public string Description
        {
            get { return "List all contacts"; }
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

            // enum contacts 
            int i = 0;
            Outlook.MAPIFolder contactFolder = outlookApplication.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            foreach (ICOMObject item in contactFolder.Items)
            {
                Outlook.ContactItem contact = item as Outlook.ContactItem;
                if (null != contact)
                {
                    i++;
                    ListViewItem listItem = listViewContacts.Items.Add(i.ToString());
                    listItem.SubItems.Add(contact.CompanyAndFullName);
                }
            }

            // close outlook and dispose
            if (!outlookApplication.FromProxyService)
                outlookApplication.Quit();
            outlookApplication.Dispose();
        }
          
        #endregion
    }
}
