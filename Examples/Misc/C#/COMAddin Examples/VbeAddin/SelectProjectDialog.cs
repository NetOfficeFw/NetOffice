using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Vbe = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using System.IO.Compression;
using System.Windows.Forms;

namespace VbeAddin
{
    public partial class SelectProjectDialog : Form
    {
        private class Project
        {
            internal Project(string name)
            {
                Name = name;
            }

            public string Name { get; private set; }

            public override string ToString()
            {
                return Name;
            }
        }

        public SelectProjectDialog(ProjectCollector projects)
        {
            InitializeComponent();
            ProjectDataGrid.SelectionChanged += delegate
            {
               ProceedButton.Enabled = ProjectDataGrid.SelectedRows.Count > 0;
            };
            ProjectDataGrid.AutoGenerateColumns = false;
            var dataSource = new List<Project>();
            foreach (var item in projects.Result)
                dataSource.Add(new Project(item));
            ProjectDataGrid.DataSource = dataSource;
        }       

        private string Selected
        {
            get
            {
                if (ProjectDataGrid.SelectedRows.Count > 0)
                    return ProjectDataGrid.SelectedRows[0].DataBoundItem.ToString();
                else
                    return null;
            }
        }

        public static string SelectProject(ProjectCollector projects)
        {
            var dialog = new SelectProjectDialog(projects);
            var dialogResult = dialog.ShowDialog();
            var selected = dialog.Selected;
            dialog.Dispose();
            if (dialogResult == DialogResult.OK)
                return selected;
            else
                return null;
        }

        private void DiscardButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void ProceedButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}