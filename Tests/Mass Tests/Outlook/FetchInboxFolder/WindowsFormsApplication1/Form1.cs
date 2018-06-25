using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        bool _cancel;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Outlook.Application application = COMObject.CreateByRunningInstance<Outlook.Application>();

            // Get inbox
            Outlook._NameSpace outlookNS = application.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            for (int i = 1; i <= 100; i++)
            {
                labelCurrentCount.Text = "Step " + i.ToString();
                Application.DoEvents();
                if (_cancel)
                    break;

                // fetch inbox
                ListInBoxFolder(inboxFolder);
            }

            labelCurrentCount.Text = "Done!";

            // close outlook and dispose
            if(application.FromProxyService)
                application.Quit();
            application.Dispose();
        }

        private void ListInBoxFolder(Outlook.MAPIFolder inboxFolder)
        {
            // setup ui
            listView1.Items.Clear();

            // we fetch the inbox folder items
            Outlook._Items items = inboxFolder.Items;
            ICOMObject item = null;
            int i = 1;
            do
            {
                if (null == item)
                {
                    item = items.GetFirst() as ICOMObject;
                    if (null == item)
                        break;
                }

                // not every item is a mail item
                Outlook.MailItem mailItem = item as Outlook.MailItem;
                if (null != mailItem)
                {
                    ListViewItem newItem = listView1.Items.Add(mailItem.SenderName);
                    newItem.SubItems.Add(mailItem.Subject);
                }

                item.Dispose();
                item = items.GetNext() as ICOMObject;
                i++;
            } while (null != item);

            // dipsose items and childs
            items.Dispose();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //_cancel = true;
        }
    }
}
