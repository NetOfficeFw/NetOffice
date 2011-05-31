using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums; 

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();
           
            // Get inbox 
            Outlook._NameSpace outlookNS = outlookApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            for (int i = 0; i < 100; i++)
            {
                // fetch inbox
                labelCurrentCount.Text = string.Format("Currently:{0}", (i + 1));
                ListInBoxFolder(outlookApplication, inboxFolder);

                this.Refresh();
            }

            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();

            labelCurrentCount.Text = "Done!";
        }

        private void ListInBoxFolder(Outlook.Application outlookApplication, Outlook.MAPIFolder inboxFolder)
        {
            // setup gui
            listView1.Items.Clear();
           
            // we fetch the inbox folder items
            Outlook._Items items = inboxFolder.Items;
            COMObject item = null;
            int i = 1;
            do
            {
                if (null == item)
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

            // dipsose items and childs
            items.Dispose();
        }
    }
}
