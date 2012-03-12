using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Outlook = NetOffice.OutlookApi;
using Office = NetOffice.OfficeApi;
using NetOffice.OutlookApi.Enums; 

namespace Example05
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

            // Create an Outlook Application object. 
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
                     ListViewItem listItem = listView1.Items.Add(i.ToString());
                     listItem.SubItems.Add(contact.CompanyAndFullName);
                 }
            }
           
            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();
        }
    }
}
