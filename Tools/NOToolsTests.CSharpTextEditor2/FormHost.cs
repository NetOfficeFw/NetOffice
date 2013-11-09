using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOToolsTests.CSharpTextEditor2
{
    public partial class FormHost : Form 
    {
        public FormHost()
        {
            InitializeComponent();

            DataHost = new DataHost();
            DataHost.Tables.Add("Persons");
            DataHost.Tables.Add("Products");

            dataDisplayControl1.OnConnect(DataHost, "Persons");
            dataDisplayControl2.OnConnect(DataHost, "Products");
            dataDisplayControl3.OnConnect(DataHost, "Persons");
            dataDisplayControl4.OnConnect(DataHost, "Products");

            buttonSaveChanges.DataBindings.Add("Enabled", dataDisplayControl1.Context, "ContainsLocalChanges", false, DataSourceUpdateMode.Never);
            buttonCancelChanges.DataBindings.Add("Enabled", dataDisplayControl1.Context, "ContainsLocalChanges", false, DataSourceUpdateMode.Never);

            buttonDoUndo.DataBindings.Add("Enabled", dataDisplayControl1.Context.Commands, "CanBackward", false, DataSourceUpdateMode.Never);
            buttonDoRedo.DataBindings.Add("Enabled", dataDisplayControl1.Context.Commands, "CanForward", false, DataSourceUpdateMode.Never);
        }

        private DataHost DataHost { get; set; }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            buttonSaveChanges.DataBindings.Remove(buttonSaveChanges.DataBindings[0]);
            buttonCancelChanges.DataBindings.Remove(buttonCancelChanges.DataBindings[0]);
            buttonDoUndo.DataBindings.Remove(buttonDoUndo.DataBindings[0]);
            buttonDoRedo.DataBindings.Remove(buttonDoRedo.DataBindings[0]);

            DataDisplayControl displayControl = e.TabPage.Controls[0] as DataDisplayControl;

            buttonSaveChanges.DataBindings.Add("Enabled", displayControl.Context, "ContainsLocalChanges", false, DataSourceUpdateMode.Never);
            buttonCancelChanges.DataBindings.Add("Enabled", displayControl.Context, "ContainsLocalChanges", false, DataSourceUpdateMode.Never);
            buttonDoUndo.DataBindings.Add("Enabled", displayControl.Context.Commands, "CanBackward", false, DataSourceUpdateMode.Never);
            buttonDoRedo.DataBindings.Add("Enabled", displayControl.Context.Commands, "CanForward", false, DataSourceUpdateMode.Never);
        }

        private void buttonCancelChanges_Click(object sender, EventArgs e)
        {
            DataDisplayControl displayControl = tabControl1.SelectedTab.Controls[0] as DataDisplayControl;
            displayControl.Context.CancelLocalChanges();
        }

        private void buttonSaveChanges_Click(object sender, EventArgs e)
        {
            DataDisplayControl displayControl = tabControl1.SelectedTab.Controls[0] as DataDisplayControl;
            displayControl.Context.ApplyLocalChanges();
        }

        private void buttonResetData_Click(object sender, EventArgs e)
        {
            DataDisplayControl displayControl = tabControl1.SelectedTab.Controls[0] as DataDisplayControl;
            displayControl.Context.ResetLocalData();
        }

        private void buttonDoUndo_Click(object sender, EventArgs e)
        {
            DataDisplayControl displayControl = tabControl1.SelectedTab.Controls[0] as DataDisplayControl;
            DataLayer.Command command = displayControl.Context.Commands.Move(false);
            if (null != command)
                command.Undo();
        }

        private void buttonDoRedo_Click(object sender, EventArgs e)
        {
            DataDisplayControl displayControl = tabControl1.SelectedTab.Controls[0] as DataDisplayControl;
            DataLayer.Command command = displayControl.Context.Commands.Move(true);
            if (null != command)
                command.Redo();
        }

        private void buttonSimulateDatabaseAction_Click(object sender, EventArgs e)
        {
            DataLayer.RootList tableProducts = DataHost.Tables["Products"];

            tableProducts[0].SetValue("Name", "ChangedName");

            tableProducts.Remove(tableProducts[1]);
            tableProducts.Remove(tableProducts[1]);

            DataLayer.RootItem item1 = tableProducts.AddNew();
            item1.SetValue("Name", "NewProduct1");

            DataLayer.RootItem item2 = tableProducts.AddNew();
            item2.SetValue("Name", "NewProduct2");

            DataLayer.RootItem item3 = tableProducts.AddNew();
            item3.SetValue("Name", "NewProduct3");
        }
    }
}
