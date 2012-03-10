using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums; 

namespace Example1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
  
        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();
 
            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

            // Get inbox 
            Outlook._NameSpace outlookNS = outlookApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            // setup gui
            listView1.Items.Clear();
            labelItemsCount.Text = string.Format("You have {0} e-mails.", inboxFolder.Items.Count);
            
            // we fetch the inbox folder items. ATTENTION: items is null if you have no items in inbox folder
            // office products initialize ALL collections on demand. this is just an example, we dont check for null here
            // NOTE: for some uninitialized collections you get an exception while accessing
            Outlook._Items items = inboxFolder.Items;
            COMObject item = null;
            int i = 1;
            do
            {
                if(null == item)
                    item = items.GetFirst();

                // not every item is a mail item
                Outlook.MailItem mailItem = item as Outlook.MailItem;
                if (null != mailItem)
                {
                    ListViewItem newItem = listView1.Items.Add(mailItem.SenderName);
                    newItem.SubItems.Add(mailItem.Subject);
                }
                item.Dispose();
                item = items.GetNext();
                i++;
            } while (null != item);

            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();
        }
    }
}
