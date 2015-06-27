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
        #region Ctor

        public Form1()
        {
            InitializeComponent();
            ShowAdminPrivileges();
            ShowRunningInstances();
        }

        #endregion

        #region Properties

        private bool IsAdministrator       
        {
            get
            {
                using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                {
                    WindowsPrincipal principal = new WindowsPrincipal(identity);
                    return principal.IsInRole(WindowsBuiltInRole.Administrator);
                    
                }                
            }
        }

        #endregion

        #region Methods

        private void ShowRunningInstances()
        {
            bool found = false;
            foreach (Process item in Process.GetProcesses())
            {
                if (item.ProcessName.Equals("excel", StringComparison.InvariantCultureIgnoreCase) ||
                   item.ProcessName.Equals("word", StringComparison.InvariantCultureIgnoreCase) ||
                    item.ProcessName.Equals("outlook", StringComparison.InvariantCultureIgnoreCase) ||
                    item.ProcessName.Equals("ppoint", StringComparison.InvariantCultureIgnoreCase) ||
                    item.ProcessName.Equals("access", StringComparison.InvariantCultureIgnoreCase) ||
                    item.ProcessName.Equals("msproject", StringComparison.InvariantCultureIgnoreCase))
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
            if (IsAdministrator)
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

        private void RunExcelTests()
        {
            NetOffice.Core.Default.Console.Name = "ExcelTests";
            NetOffice.Core.Default.Console.EnableSharedOutput = true;

            ExcelTestsVB.TestAssembly excelVB = new ExcelTestsVB.TestAssembly();
            foreach (ITestPackage item in excelVB.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }

            ExcelTestsCSharp.TestAssembly excelCSharp = new ExcelTestsCSharp.TestAssembly();
            foreach (ITestPackage item in excelCSharp.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }
        }

        private void RunOutlookTests()
        {
            NetOffice.Core.Default.Console.Name = "OutlookTests";
            NetOffice.Core.Default.Console.EnableSharedOutput = true;

            OutlookTestsVB.TestAssembly outlookVB = new OutlookTestsVB.TestAssembly();
            foreach (ITestPackage item in outlookVB.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }

            OutlookTestsCSharp.TestAssembly outlookCSharp = new OutlookTestsCSharp.TestAssembly();
            foreach (ITestPackage item in outlookCSharp.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }
        }

        private void RunWordTests()
        {
            NetOffice.Core.Default.Console.Name = "WordTests";
            NetOffice.Core.Default.Console.EnableSharedOutput = true;

            WordTestsVB.TestAssembly wordVB = new WordTestsVB.TestAssembly();
            foreach (ITestPackage item in wordVB.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }

            WordTestsCSharp.TestAssembly wordCSharp = new WordTestsCSharp.TestAssembly();
            foreach (ITestPackage item in wordCSharp.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }
        }

        private void RunPowerPointTests()
        {
            NetOffice.Core.Default.Console.Name = "PowerPointTests";
            NetOffice.Core.Default.Console.EnableSharedOutput = true;

            PowerPointTestsVB.TestAssembly powerVB = new PowerPointTestsVB.TestAssembly();
            foreach (ITestPackage item in powerVB.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }

            PowerPointTestsCSharp.TestAssembly powerCSharp = new PowerPointTestsCSharp.TestAssembly();
            foreach (ITestPackage item in powerCSharp.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }
        }

        private void RunAccessTests()
        {
            NetOffice.Core.Default.Console.Name = "AccessTests";
            NetOffice.Core.Default.Console.EnableSharedOutput = true;

            AccessTestsVB.TestAssembly accessVB = new AccessTestsVB.TestAssembly();
            foreach (ITestPackage item in accessVB.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }

            AccessTestsCSharp.TestAssembly accessCSharp = new AccessTestsCSharp.TestAssembly();
            foreach (ITestPackage item in accessCSharp.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }
        }

        private void RunProjectTests()
        {
            NetOffice.Core.Default.Console.Name = "ProjectTests";
            NetOffice.Core.Default.Console.EnableSharedOutput = true;

            ProjectTestsVB.TestAssembly projectVB = new ProjectTestsVB.TestAssembly();
            foreach (ITestPackage item in projectVB.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }

            ProjectTestsCSharp.TestAssembly projectCSharp = new ProjectTestsCSharp.TestAssembly();
            foreach (ITestPackage item in projectCSharp.LoadTestPackages())
            {
                ShowCurrentTestPackge(item);
                TestResult result = null;
                try
                {
                    result = item.DoTest();
                }
                catch (Exception exception)
                {
                    result = new TestResult(false, TimeSpan.MinValue, "Unexpected Error.", exception, "");
                }
                AddTestResult(item, result);
            }
        }

        private void AddTestResult(ITestPackage package, TestResult result)
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

        private void ShowCurrentTestPackge(ITestPackage package)
        {
            labelCurrentTest.Text = "Do Test: " + string.Format("{0} {1} {2} {3}", package.OfficeProduct, package.Name, package.Language, package.Description);
            labelCurrentTest.Refresh();
        }
       
        #endregion

        #region Trigger

        private void listViewResults_DoubleClick(object sender, EventArgs e)
        {
            if ((listViewResults.SelectedItems.Count > 0) && (listViewResults.SelectedItems[0].Tag is Exception))
            {
                Exception exception = listViewResults.SelectedItems[0].Tag as Exception;
                ExceptionDialog dialog = new ExceptionDialog(exception);
                dialog.ShowDialog(this);
            }
        }

        private void labelRunningInstances_Click(object sender, EventArgs e)
        {
            ShowRunningInstances();
        }

        private void buttonTest_Click(object sender, EventArgs e)
        {
            try
            {   
                buttonTest.Enabled =false;
                listViewResults.Items.Clear();

                if (checkBoxExcel.Checked)
                    RunExcelTests();

                if (checkBoxOutlook.Checked)
                    RunOutlookTests();

                if (checkBoxWord.Checked)
                    RunWordTests();

                if (checkBoxPowerPoint.Checked)
                    RunPowerPointTests();

                if (checkBoxAccess.Checked)
                    RunAccessTests();

                if (checkBoxProject.Checked)
                    RunProjectTests();
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

        #endregion
    }
}
