using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Tests.Core;

namespace Host
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void buttonTest_Click(object sender, EventArgs e)
        {
            try
            {   
                buttonTest.Enabled =false;
                listViewResults.Items.Clear();
                
                #region Excel

                ExcelTestsVB.TestAssembly excelVB = new ExcelTestsVB.TestAssembly();
                foreach (ITestPackage item in excelVB.LoadTestPackages())
                {
                    ShowCurrent(item);
                    TestResult result = item.DoTest();
                    AddResult(item, result);
                }

                ExcelTestsCSharp.TestAssembly excelCSharp = new ExcelTestsCSharp.TestAssembly();
                foreach (ITestPackage item in excelCSharp.LoadTestPackages())
                {
                    ShowCurrent(item);
                    TestResult result = item.DoTest();
                    AddResult(item, result);
                }

                #endregion

            }
            catch (Exception exception)
            {
                string message = Environment.NewLine + Environment.NewLine + exception.Message;
                MessageBox.Show("We have an Error" + message, "Huston", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelCurrentTest.Text = string.Empty;
                buttonTest.Enabled = true;
            }
        }

        private void AddResult(ITestPackage package, TestResult result)
        {
            if (result.Sucseed)
            {
                ListViewItem newItem = listViewResults.Items.Add(string.Format("{0} {1} {2}", package.OfficeProduct, package.Name, package.Language));
                newItem.SubItems.Add(string.Format("Test passed in {0}.", result.TimeElapsed));
                newItem.StateImageIndex = 0;
            }
            else
            {
                ListViewItem newItem = listViewResults.Items.Add(string.Format("{0} {1} {2}", package.OfficeProduct, package.Name, package.Language));
                newItem.SubItems.Add(string.Format("Test failed. Message:{0}.", result.ErrorInfo));
                newItem.Tag = result.Exception;
                newItem.StateImageIndex = 1;
            }
            listViewResults.Refresh();
        }

        private void ShowCurrent(ITestPackage package)
        {
            labelCurrentTest.Text = "Do Test: " + string.Format("{0} {1} {2} {3}", package.OfficeProduct, package.Name, package.Language, package.Description);
            labelCurrentTest.Refresh();
        }

        private void listViewResults_DoubleClick(object sender, EventArgs e)
        {
            if( (listViewResults.SelectedItems.Count > 0) && (listViewResults.SelectedItems[0].Tag is Exception) )
            {
                Exception exception = listViewResults.SelectedItems[0].Tag as Exception;
                ExceptionDialog dialog = new ExceptionDialog(exception);
                dialog.ShowDialog(this);
            }
        }
    }
}
