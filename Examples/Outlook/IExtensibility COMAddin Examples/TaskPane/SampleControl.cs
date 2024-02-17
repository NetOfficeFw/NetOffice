﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Extensibility;

using NetOffice;
using Office = NetOffice.OfficeApi;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;


namespace COMAddinTaskPaneExampleCS4
{
    public partial class SampleControl : UserControl
    {
        List<Customer> _customers;

        public SampleControl()
        {
            InitializeComponent();
            LoadSampleCustomerData();
            UpdateSearchResult();
        }

        #region Private Methods

        private void LoadSampleCustomerData()
        {
            _customers = new List<Customer>();

            string embeddedCustomerXmlContent = ReadString("SampleData.CustomerData.xml");
            XmlDocument document = new XmlDocument();
            document.LoadXml(embeddedCustomerXmlContent);
            foreach (XmlNode customerNode in document.DocumentElement.ChildNodes)
            {
                int id = Convert.ToInt32(customerNode.Attributes["ID"].Value);
                string name = customerNode.Attributes["Name"].Value;
                string company = customerNode.Attributes["Company"].Value;
                string mail = customerNode.Attributes["Mail"].Value;
                string city = customerNode.Attributes["City"].Value;
                string postalCode = customerNode.Attributes["PostalCode"].Value;
                string country = customerNode.Attributes["Country"].Value;
                string phone = customerNode.Attributes["Phone"].Value;

                _customers.Add(new Customer(id, name, company, mail, city, postalCode, country, phone));
            }
        }

        private string ReadString(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            System.IO.StreamReader textStreamReader = null;
            try
            {
                Assembly assembly = typeof(Addin).Assembly;
                ressourceStream = assembly.GetManifestResourceStream(assembly.GetName().Name + "." + ressourcePath);
                if (ressourceStream == null)
                    throw (new System.IO.IOException("Error accessing resource Stream."));

                textStreamReader = new System.IO.StreamReader(ressourceStream);
                if (textStreamReader == null)
                    throw (new System.IO.IOException("Error accessing resource File."));

                string text = textStreamReader.ReadToEnd();
                return text;
            }
            catch (Exception exception)
            {
                throw (exception);
            }
            finally
            {
                if (null != textStreamReader)
                    textStreamReader.Close();
                if (null != ressourceStream)
                    ressourceStream.Close();
            }
        }

        private void UpdateSearchResult()
        {
            listViewSearchResults.Items.Clear();
            foreach (Customer item in _customers)
            {
                if (item.Name.IndexOf(textBoxSearch.Text.Trim(), StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    ListViewItem viewItem = listViewSearchResults.Items.Add("");
                    viewItem.SubItems.Add(item.ID.ToString());
                    viewItem.SubItems.Add(item.Name);
                    viewItem.ImageIndex = 0;
                    viewItem.Tag = item;
                }
            }
        }

        private void UpdateDetails()
        {
            if (listViewSearchResults.SelectedItems.Count > 0)
            {
                Customer selectedCustomer = listViewSearchResults.SelectedItems[0].Tag as Customer;
                propertyGridDetails.SelectedObject = selectedCustomer;
            }
            else
                propertyGridDetails.SelectedObject = null;
        }

        #endregion

        #region UI Trigger

        private void listViewSearchResults_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (listViewSearchResults.SelectedItems.Count > 0)
                {
                    Customer selectedCustomer = listViewSearchResults.SelectedItems[0].Tag as Customer;

                    // create MailItem
                    Outlook.MailItem mailItem = Addin.Application.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;                    
                    mailItem.BodyFormat = OlBodyFormat.olFormatRichText;
                    mailItem.To = selectedCustomer.Mail;
                    mailItem.Body = "Hello " + selectedCustomer.Name;
                    mailItem.Subject = "<Insert Subject>";
                    mailItem.Display();
                    mailItem.Dispose();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "An error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void listViewSearchResults_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            try
            {
                UpdateDetails();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "An error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                UpdateSearchResult();
                UpdateDetails();
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "An error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        #endregion
    }
}
