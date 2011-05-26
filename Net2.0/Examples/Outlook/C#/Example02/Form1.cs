using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums; 

namespace Example02
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

            // Create a new TaskItem.
            Outlook.TaskItem newTask = outlookApplication.CreateItem(OlItemType.olTaskItem) as Outlook.TaskItem;

            // Configure the task at hand and save it.
            newTask.Subject = "Don't forget to check for NetOffice.DeveloperUtils updates";
            newTask.Body = "check updates here: http://netoffice.codeplex.com";
            newTask.DueDate = DateTime.Now;
            newTask.Importance = OlImportance.olImportanceHigh;
            
            newTask.Save();
           
            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();

            MessageBox.Show(this,"Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
