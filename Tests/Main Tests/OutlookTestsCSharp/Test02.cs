using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookTestsCSharp
{
    public class Test02 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test02"; }
        }

        public string Description
        {
            get { return "Create a task item."; }
        }

        public string OfficeProduct
        {
            get { return "Outlook"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Outlook.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                // start outlook
                application = new Outlook.Application();
                NetOffice.OutlookSecurity.Suppress.Enabled = true;

                Outlook.TaskItem newTask = application.CreateItem(OlItemType.olTaskItem) as Outlook.TaskItem;
                newTask.Subject = "Test item";
                newTask.Body = "hello";
                newTask.DueDate = DateTime.Now;
                newTask.Importance = OlImportance.olImportanceHigh;
                newTask.Close(OlInspectorClose.olDiscard);

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    application.Quit();
                    application.Dispose();
                }
            }
        }

        #endregion
    }
}
