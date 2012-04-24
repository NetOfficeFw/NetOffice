using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;

using NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

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
                int    id = Convert.ToInt32(customerNode.Attributes["ID"].Value);
                string name = customerNode.Attributes["Name"].Value;
                string company = customerNode.Attributes["Company"].Value;
                string city = customerNode.Attributes["City"].Value;
                string postalCode = customerNode.Attributes["PostalCode"].Value;
                string country = customerNode.Attributes["Country"].Value;
                string phone = customerNode.Attributes["Phone"].Value;

                _customers.Add(new Customer(id, name, company, city, postalCode, country, phone));                
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

        public static string ToRangeAddress(int rowIndex, int columnIndex)
        {
            if (columnIndex < 1) throw (new ArgumentOutOfRangeException("Invalid Argument. columnIndex must be > 0"));
            if (rowIndex < 1) throw (new ArgumentOutOfRangeException("Invalid Argument. rowIndex must be > 0"));

            string[] columnChars = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            if (columnIndex <= columnChars.Length)
                return columnChars[columnIndex - 1] + rowIndex.ToString();

            int multi = columnIndex / columnChars.Length;
            string pre = columnChars[multi - 1];

            int newx = columnIndex;
            newx -= (multi * columnChars.Length);
            return pre + columnChars[newx - 1] + rowIndex.ToString();
        }

        private string CalculateRangeArea(int rowIndex, int columnIndex, int countOfProperties)
        {
            string startRangeAddress = ToRangeAddress(rowIndex, columnIndex);
            string endEndRangeAddress = ToRangeAddress(rowIndex + countOfProperties - 1, columnIndex + 1);
            return startRangeAddress + ":" + endEndRangeAddress;
        }

        private object[,] ToStringArray(Customer customer)
        {
            object[,] customerPropertiesArray = new object[7, 2];

            customerPropertiesArray[0, 0] = "ID:";
            customerPropertiesArray[0, 1] = customer.ID.ToString();

            customerPropertiesArray[1, 0] = "Name:";
            customerPropertiesArray[1, 1] = customer.Name;

            customerPropertiesArray[2, 0] = "Company:";
            customerPropertiesArray[2, 1] = customer.Company;

            customerPropertiesArray[3, 0] = "City:";
            customerPropertiesArray[3, 1] = customer.City;

            customerPropertiesArray[4, 0] = "Postal Code:";
            customerPropertiesArray[4, 1] = customer.PostalCode;

            customerPropertiesArray[5, 0] = "Country:";
            customerPropertiesArray[5, 1] = customer.Country;

            customerPropertiesArray[6, 0] = "Phone:";
            customerPropertiesArray[6, 1] = customer.Phone;

            return customerPropertiesArray;
        }
       
        #endregion

        #region UI Trigger

        private void listViewSearchResults_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (listViewSearchResults.SelectedItems.Count > 0)
                {
                    Excel.Worksheet activeSheet = Addin.Application.ActiveSheet as Excel.Worksheet;
                    Excel.Range activeCell = Addin.Application.ActiveCell;
                    if (null != activeCell)
                    {
                        int rowIndex = activeCell.Row;
                        int columnIndex = activeCell.Column;

                        string targetRangeAddress = CalculateRangeArea(rowIndex, columnIndex, 7);

                        Customer selectedCustomer = listViewSearchResults.SelectedItems[0].Tag as Customer;

                        Excel.Range targetRange = activeSheet.Range(targetRangeAddress);
                        targetRange.Value2 = ToStringArray(selectedCustomer);
                        targetRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        activeSheet.Columns[targetRange.Column].AutoFit();

                        activeCell.Dispose();
                        activeSheet.Dispose();
                    }
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
