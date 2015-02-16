using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using ExampleBase;

using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookExamplesCS4
{
    /// <summary>
    /// Example 4 - All Contacts
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
            get { return HostApplication.LCID == 1033 ? "Example05" : "Beispiel05"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "List all contacts" : "Alle Kontakte auflisten"; }
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
            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

            // enum contacts 
            int i = 0;
            Outlook.MAPIFolder contactFolder = outlookApplication.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            foreach (COMObject item in contactFolder.Items)
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
            outlookApplication.Quit();
            outlookApplication.Dispose();
        }
          
        #endregion
    }
}
