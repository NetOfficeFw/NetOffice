using System;
using ExampleBase;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookExamplesCS4
{
    /// <summary>
    /// Example 2 - Create Task Item
    /// </summary>
    internal class Example02 :IExample
    {
        public void RunExample()
        {
            // start outlook by trying to access running application first
            Outlook.Application outlookApplication = new Outlook.Application(true);

            // create a new TaskItem.
            Outlook.TaskItem newTask = outlookApplication.CreateItem(OlItemType.olTaskItem) as Outlook.TaskItem;

            // Configure the task at hand and save it.
            newTask.Subject = "Don't forget to check for NoScript updates";
            newTask.Body = "check updates here: https://addons.mozilla.org/de/firefox/addon/noscript";
            newTask.DueDate = DateTime.Now;
            newTask.Importance = OlImportance.olImportanceHigh;
            newTask.Save();

            // close outlook and dispose
            if (!outlookApplication.FromProxyService)
                outlookApplication.Quit();
            outlookApplication.Dispose();

            HostApplication.ShowFinishDialog("Done!", null);
        }

        public string Caption
        {
            get { return  "Example02"; }
        }

        public string Description
        {
            get { return "Create task item"; }
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }
        
        public System.Windows.Forms.UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
