using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tests.Core;
using MSProject = NetOffice.MSProjectApi;
using NetOffice.MSProjectApi.Enums;

namespace ProjectTestsCSharp
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
            get { return "Test events"; }
        }

        public string OfficeProduct
        {
            get { return "Project"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            NetOffice.Settings.Default.MessageFilter.Enabled = true;
            MSProject.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                application = new MSProject.Application();

                application.ProjectTaskNewEvent += new MSProject.Application_ProjectTaskNewEventHandler(ApplicationProjectTaskNewEvent);
                application.ProjectBeforeCloseEvent += new MSProject.Application_ProjectBeforeCloseEventHandler(ApplicationProjectBeforeCloseEvent);
                application.ProjectBeforeTaskChangeEvent += new MSProject.Application_ProjectBeforeTaskChangeEventHandler(ApplicationProjectBeforeTaskChangeEvent);
                application.ProjectBeforeTaskDeleteEvent += new MSProject.Application_ProjectBeforeTaskDeleteEventHandler(ApplicationProjectBeforeTaskDeleteEvent);
               
                MSProject.Project newProject = application.Projects.Add();
                MSProject.Task task1 = newProject.Tasks.Add("Task 1");
                MSProject.Task task2 =  newProject.Tasks.Add("Task 2");
                
                task2.Delete();

                application.FileCloseAll(PjSaveType.pjDoNotSave);

                if (TaskDeleteEventCalled && TaskChangeEventCalled && BeforeCloseEventCalled && TaskChangeEventCalled)
                    return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
                else 
                {
                    string errorMessage = "";
                    if (!TaskDeleteEventCalled)
                        errorMessage += "ProjectBeforeTaskDeleteEvent failed ";
                    if (!TaskChangeEventCalled)
                        errorMessage += "ProjectBeforeTaskChangeEvent failed ";
                    if (!BeforeCloseEventCalled)
                        errorMessage += "ProjectBeforeCloseEvent failed ";
                    if (!TaskNewEventCalled)
                        errorMessage += "ProjectTaskNewEvent failed ";

                    return new TestResult(false, DateTime.Now.Subtract(startTime), errorMessage, null, "");
                }
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                NetOffice.Settings.Default.MessageFilter.Enabled = false;

                if (null != application)
                {
                    application.Quit(PjSaveType.pjDoNotSave);
                    application.Dispose();
                }
            }
        }

        #endregion

        private bool TaskDeleteEventCalled { get; set; }
        private bool TaskChangeEventCalled { get; set; }
        private bool BeforeCloseEventCalled { get; set; }
        private bool TaskNewEventCalled { get; set; }


        void ApplicationProjectBeforeTaskDeleteEvent(MSProject.Task tsk, ref bool Cancel)
        {
            TaskDeleteEventCalled = true;
        }

        void ApplicationProjectBeforeTaskChangeEvent(MSProject.Task tsk, PjField Field, object NewVal, ref bool Cancel)
        {
            TaskChangeEventCalled = true;
        }

        void ApplicationProjectBeforeCloseEvent(MSProject.Project pj, ref bool Cancel)
        {
            BeforeCloseEventCalled = true;
        }

        void ApplicationProjectTaskNewEvent(MSProject.Project pj, int ID)
        {
            TaskNewEventCalled = true;
        }
    }
}
