using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Security.Principal;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using Tests.Core;

namespace Host
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ShowAdminPrivileges();
            CheckRunningInstances();
        }
        
        private void CheckRunningInstances()
        {
            bool found = false;
            foreach (Process item in Process.GetProcesses())
            {
                if (item.ProcessName.Equals("excel", StringComparison.InvariantCultureIgnoreCase))
                {
                    found = true;
                    break;
                }
            }
            if (found)
                labelRunningInstances.Visible = true;
        }

        private void ShowAdminPrivileges()
        {
            if (IsAdministrator())
            {
                labelAdmin.ForeColor = Color.Green;
                labelAdmin.Text = "Program runs with Administrator Privileges.";
            }
            else
            {
                labelAdmin.ForeColor = Color.Red;
                labelAdmin.Text = "Program runs without Administrator Privileges.";
            }
        }

        private bool IsAdministrator()
        {
           WindowsIdentity identity =  WindowsIdentity.GetCurrent();
           WindowsPrincipal principal = new WindowsPrincipal (identity);
           return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
       
        private void AddResult(ITestPackage package, TestResult result)
        {
            if (result.Sucseed)
            {
                ListViewItem newItem = listViewResults.Items.Add(string.Format("{0} {1} {2}", package.OfficeProduct, package.Name, package.Language));
                newItem.SubItems.Add(string.Format("Test passed in {0}. {1}", result.TimeElapsed, result.Hints));
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

        void buttonTest_Click(object sender, EventArgs e)
        {
            try
            {   
                buttonTest.Enabled =false;
                listViewResults.Items.Clear();
                
                #region Excel

                if (checkBoxExcel.Checked)
                { 
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
                }

                #endregion

                #region Outlook

                if (checkBoxOutlook.Checked)
                { 
                    OutlookTestsVB.TestAssembly outlookVB = new OutlookTestsVB.TestAssembly();
                    foreach (ITestPackage item in outlookVB.LoadTestPackages())
                    {
                        ShowCurrent(item);
                        TestResult result = item.DoTest();
                        AddResult(item, result);
                    }

                    OutlookTestsCSharp.TestAssembly outlookCSharp = new OutlookTestsCSharp.TestAssembly();
                    foreach (ITestPackage item in outlookCSharp.LoadTestPackages())
                    {
                        ShowCurrent(item);
                        TestResult result = item.DoTest();
                        AddResult(item, result);
                    }
                }

                #endregion

                #region Word

                if (checkBoxWord.Checked)
                {
                    WordTestsVB.TestAssembly wordVB = new WordTestsVB.TestAssembly();
                    foreach (ITestPackage item in wordVB.LoadTestPackages())
                    {
                        ShowCurrent(item);
                        TestResult result = item.DoTest();
                        AddResult(item, result);
                    }

                    WordTestsCSharp.TestAssembly wordCSharp = new WordTestsCSharp.TestAssembly();
                    foreach (ITestPackage item in wordCSharp.LoadTestPackages())
                    {
                        ShowCurrent(item);
                        TestResult result = item.DoTest();
                        AddResult(item, result);
                    }
                }

                #endregion

                #region PowerPoint

                if (checkBoxPowerPoint.Checked)
                {
                    PowerPointTestsVB.TestAssembly powerVB = new PowerPointTestsVB.TestAssembly();
                    foreach (ITestPackage item in powerVB.LoadTestPackages())
                    {
                        ShowCurrent(item);
                        TestResult result = item.DoTest();
                        AddResult(item, result);
                    }

                    PowerPointTestsCSharp.TestAssembly powerCSharp = new PowerPointTestsCSharp.TestAssembly();
                    foreach (ITestPackage item in powerCSharp.LoadTestPackages())
                    {
                        ShowCurrent(item);
                        TestResult result = item.DoTest();
                        AddResult(item, result);
                    }
                }

                #endregion

                #region AccessPoint

                if (checkBoxAccess.Checked)
                {
                    AccessTestsVB.TestAssembly accessVB = new AccessTestsVB.TestAssembly();
                    foreach (ITestPackage item in accessVB.LoadTestPackages())
                    {
                        ShowCurrent(item);
                        TestResult result = item.DoTest();
                        AddResult(item, result);
                    }
                }

                if (checkBoxAccess.Checked)
                {
                    AccessTestsCSharp.TestAssembly accessCSharp = new AccessTestsCSharp.TestAssembly();
                    foreach (ITestPackage item in accessCSharp.LoadTestPackages())
                    {
                        ShowCurrent(item);
                        TestResult result = item.DoTest();
                        AddResult(item, result);
                    }
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
