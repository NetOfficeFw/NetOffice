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
    public partial class Example01 : UserControl , IExample 
    {
        IHost _hostApplication;

        public Example01()
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
            get { return _hostApplication.LCID == 1033 ? "Example01" : "Beispiel01"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Inbox Folder" : "Posteingang"; }
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

            // get inbox 
            Outlook._NameSpace outlookNS = outlookApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            // setup gui
            listViewInboxFolder.Items.Clear();
            labelItemsCount.Text = string.Format("You have {0} e-mails.", inboxFolder.Items.Count);

            // we fetch the inbox folder items.
            foreach (COMObject item in inboxFolder.Items)
            {
                // not every item in the inbox is a mail item
                Outlook.MailItem mailItem = item as Outlook.MailItem;
                if (null != mailItem)
                {
                    ListViewItem newItem = listViewInboxFolder.Items.Add(mailItem.SenderName);
                    newItem.SubItems.Add(mailItem.Subject);
                }
            }

            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();
        }
        
        #endregion
    }
}
