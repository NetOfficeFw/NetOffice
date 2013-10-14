using System;
using System.Drawing;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NOTools.CSharpTextEditor.GACManagedAccess;

namespace NOTools.CSharpTextEditor
{
    public partial class ReferencesDialog : Form
    {
        private bool _lastItemInvert = true;

        public ReferencesDialog()
        {
            InitializeComponent();
            GACAssemblyItems = new List<ListViewItem>();
            GACAssemblies.BeginLoadAssemblyInformations(UpdateTreeView, 20);
        }

        internal List<ListViewItem> GACAssemblyItems { get; private set; }
        internal bool ReadyForAction { get; private set; }

        private void UpdateTreeView(GACAssembly[] gacAssemblies, bool isLastUpdate)
        { 
            if(listView1.InvokeRequired)
            {
                listView1.Invoke(new LoadAssemblyUpdateEventHandler(UpdateTreeView), new object[] { gacAssemblies, isLastUpdate });
            }
            else
            {
                if (listView1.Items[0].Text == "Please wait...")
                    listView1.Items.Clear();
                
                ListViewItem[] viewItems = new ListViewItem[gacAssemblies.Count()];
                for (int i = 0; i < gacAssemblies.Count(); i++)
                {
                    viewItems[i] = new ListViewItem();
                    viewItems[i].Text = gacAssemblies[i].Name;
                    viewItems[i].SubItems.Add(gacAssemblies[i].Version.ToString());
                    viewItems[i].SubItems.Add(gacAssemblies[i].PublicKeyToken);
                    viewItems[i].SubItems.Add(gacAssemblies[i].Path);
                    viewItems[i].Font = new System.Drawing.Font("Arial", 10);
                    viewItems[i].BackColor = _lastItemInvert == true ? Color.White : Color.LavenderBlush;
                    if (!CheckFilter(gacAssemblies[i].Name))
                        continue;
                    _lastItemInvert = !_lastItemInvert;
                    GACAssemblyItems.Add(viewItems[i]);
                }
                listView1.Items.AddRange(viewItems);
            }

            ReadyForAction = isLastUpdate;
        }

        private void RefreshAssemblyFilter()
        {
            listView1.Items.Clear();
            foreach (ListViewItem item in GACAssemblyItems)
            {
                if (CheckFilter(item.Text))
                    listView1.Items.Add(item);
            }
        }

        private bool CheckFilter(string name)
        {
            if (String.IsNullOrWhiteSpace(textBoxNameFilter.Text))
                return true;

            int position = name.IndexOf(textBoxNameFilter.Text, StringComparison.InvariantCultureIgnoreCase);
            return position >= 0;
        }
        
        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (!ReadyForAction)
                return;
            GACAssemblyItems.Sort(new ListViewGacItemComparer(e.Column));
            listView1.Items.Clear();
            listView1.Items.AddRange(GACAssemblyItems.ToArray());
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPageGac)
                buttonOk.Enabled = listView1.SelectedItems.Count > 0;     
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void textBoxNameFilter_KeyDown(object sender, KeyEventArgs e)
        {
            if (false == ReadyForAction || e.KeyCode != Keys.Return)
                return;
            RefreshAssemblyFilter();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPageGac)
                buttonOk.Enabled = listView1.SelectedItems.Count > 0;
            else
                buttonOk.Enabled = false;
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            if (!ReadyForAction)
                return;
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (!ReadyForAction)
                return;
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            if (!ReadyForAction)
                return;
        }
    }
}
