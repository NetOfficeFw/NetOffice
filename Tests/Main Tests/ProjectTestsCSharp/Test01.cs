using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tests.Core;
using MSProject = NetOffice.MSProjectApi;
using NetOffice.MSProjectApi.Enums;

namespace ProjectTestsCSharp
{
    public class Test01 : ITestPackage
    {   
        #region TestPackage Member
         
        public string Name
        {
            get { return "Test01"; }
        }

        public string Description
        {
            get { return "Add a new project and add 10 tasks."; }
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
            MSProject.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                application = new MSProject.Application();

                MSProject.Project newProject = application.Projects.Add();
                
                newProject.Tasks.Add("Task 0");
                newProject.Tasks.Add("Task 1");
                newProject.Tasks.Add("Task 2");
                newProject.Tasks.Add("Task 3");
                newProject.Tasks.Add("Task 4");
                newProject.Tasks.Add("Task 5");
                newProject.Tasks.Add("Task 6");
                newProject.Tasks.Add("Task 7");
                newProject.Tasks.Add("Task 8");
                newProject.Tasks.Add("Task 9");

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
                    application.Quit(PjSaveType.pjDoNotSave);
                    application.Dispose();
                }
            }
        }

        #endregion
    }
}
