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
            textBoxFolder.Text = Application.StartupPath;
        }

        private void buttonSelectFolder_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                if (DialogResult.OK == dialog.ShowDialog(this))
                {
                    textBoxFolder.Text = dialog.SelectedPath;
                }
            }
            catch (Exception exception)
            {
                string message = Environment.NewLine + Environment.NewLine + exception.Message;
                MessageBox.Show("We have an Error" + message, "Huston", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonOpenFolder_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(textBoxFolder.Text);
            }
            catch (Exception exception)
            {
                string message = Environment.NewLine + Environment.NewLine + exception.Message;
                MessageBox.Show("We have an Error" + message, "Huston", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddResult(string name, bool result)
        {
            if (result)
            {
                ListViewItem newItem = listViewResults.Items.Add(name);
                newItem.SubItems.Add("Test passed.");
                newItem.StateImageIndex = 0;
            }
            else
            {
                ListViewItem newItem = listViewResults.Items.Add(name);
                newItem.SubItems.Add("Test failed.");
                newItem.StateImageIndex = 1;
            }
            listViewResults.Refresh();
        }

        private void ShowCurrent(string name)
        {
            labelCurrentTest.Text = "Do Test: " + name;
            labelCurrentTest.Refresh();
        }

        private void buttonTest_Click(object sender, EventArgs e)
        {
            try
            {   
                buttonTest.Enabled =false;

                if (!System.IO.Directory.Exists(textBoxFolder.Text))
                    System.IO.Directory.CreateDirectory(textBoxFolder.Text);
                else
                {
                    foreach (string item in System.IO.Directory.GetFiles(textBoxFolder.Text, "*.log"))
                        System.IO.File.Delete(item);
                }

                #region Excel
                
                ShowCurrent("Excel Test01");
                ITestPackage excelTest01 = new ExcelTests.Test01();
                AddResult("Excel Test01", excelTest01.DoTest(textBoxFolder.Text));

                ShowCurrent("Excel Test02");
                ITestPackage excelTest02 = new ExcelTests.Test02();
                AddResult("Excel Test02", excelTest02.DoTest(textBoxFolder.Text));

                ShowCurrent("Excel Test03");
                ITestPackage excelTest03 = new ExcelTests.Test03();
                AddResult("Excel Test03", excelTest03.DoTest(textBoxFolder.Text));

                ShowCurrent("Excel Test04");
                ITestPackage excelTest04 = new ExcelTests.Test04();
                AddResult("Excel Test04", excelTest04.DoTest(textBoxFolder.Text));

                ShowCurrent("Excel Test05");
                ITestPackage excelTest05 = new ExcelTests.Test05();
                AddResult("Excel Test05", excelTest05.DoTest(textBoxFolder.Text));

                ShowCurrent("Excel Test06");
                ITestPackage excelTest06 = new ExcelTests.Test06();
                AddResult("Excel Test06", excelTest06.DoTest(textBoxFolder.Text));
              
                ShowCurrent("Excel Test07");
                ITestPackage excelTest07 = new ExcelTests.Test07();
                AddResult("Excel Test07", excelTest07.DoTest(textBoxFolder.Text));

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
                buttonTest.Enabled =true;
            }
        }

        private void buttonSetDefaultFolder_Click(object sender, EventArgs e)
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            textBoxFolder.Text = System.IO.Path.Combine(localAppData, "!NetOfficeTest");
        }
 
    }
}
