using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        MSProject.Application app;

        public Form1()
        {
            InitializeComponent();
            app = new MSProject.Application();
            app.ProjectBeforeTaskNew += new MSProject._EProjectApp2_ProjectBeforeTaskNewEventHandler(app_ProjectBeforeTaskNew);

            app.NewProject += new MSProject._EProjectApp2_NewProjectEventHandler(app_NewProject);
            app.ProjectBeforeClose += new MSProject._EProjectApp2_ProjectBeforeCloseEventHandler(app_ProjectBeforeClose);
            app.ProjectBeforeTaskChange += new MSProject._EProjectApp2_ProjectBeforeTaskChangeEventHandler(app_ProjectBeforeTaskChange);
            app.ProjectBeforeTaskDelete += new MSProject._EProjectApp2_ProjectBeforeTaskDeleteEventHandler(app_ProjectBeforeTaskDelete);
            MSProject.Project newProject = app.Projects.Add();
            newProject.Tasks.Add("Task 1");
            newProject.Tasks.Add("Task 2");
        }

        void app_ProjectBeforeTaskDelete(MSProject.Task tsk, ref bool Cancel)
        {
             
        }

        void app_ProjectBeforeTaskChange(MSProject.Task tsk, MSProject.PjField Field, object NewVal, ref bool Cancel)
        {
           
        }

        void app_ProjectBeforeClose(MSProject.Project pj, ref bool Cancel)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MSProject.Project proj = app.Projects.Add();
            proj.Tasks.Add(Type.Missing, Type.Missing);
            app.Quit(MSProject.PjSaveType.pjDoNotSave);
        }

        void app_ProjectBeforeTaskNew(MSProject.Project pj, ref bool Cancel)
        {
          
        }

        void app_NewProject(MSProject.Project pj)
        {
           
        }
    }
}
