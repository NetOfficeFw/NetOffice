using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using ExampleBase;

using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookExamplesCS4
{
    class Example02 :IExample
    {
        IHost _hostApplication;

        #region IExample Member

        public void RunExample()
        {
            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

            // create a new TaskItem.
            Outlook.TaskItem newTask = outlookApplication.CreateItem(OlItemType.olTaskItem) as Outlook.TaskItem;

            // Configure the task at hand and save it.
            newTask.Subject = "Don't forget to check for NetOffice.DeveloperToolbox updates";
            newTask.Body = "check updates here: http://netoffice.codeplex.com/releases";
            newTask.DueDate = DateTime.Now;
            newTask.Importance = OlImportance.olImportanceHigh;
            newTask.Save();

            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();

            _hostApplication.ShowFinishDialog("Done!", null);
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example02" : "Beispiel02"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Create task item" : "Ein TaskItem erstellen"; }
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }
        
        public System.Windows.Forms.UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
