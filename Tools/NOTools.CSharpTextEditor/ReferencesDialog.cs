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
        #region Fields

        private bool _lastItemInvert = true;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ReferencesDialog()
        {
            InitializeComponent();
            GACAssemblyItems = new List<ListViewItem>();
            GACAssemblies.BeginLoadAssemblyInformations(UpdateGACListView, 20);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Use selected assemblies
        /// </summary>
        public SelectedAssembly[] SelectedAssemblies
        {
            get
            {
                List<SelectedAssembly> list = new List<SelectedAssembly>();
                if (tabControl1.SelectedTab == tabPageGac)
                {
                    foreach (ListViewItem item in listView1.SelectedItems)
                        list.Add(new SelectedAssembly(item.Text, item.SubItems[3].Text));
                }
                else 
                {
                    foreach (var item in openFilePanel1.SelectedFiles)
                    {
                        string name = System.IO.Path.GetFileNameWithoutExtension(item);
                        list.Add(new SelectedAssembly(name, item));
                    }
                }
                return list.ToArray();
            }
        }

        /// <summary>
        /// All collected GAC assemblies
        /// </summary>
        internal List<ListViewItem> GACAssemblyItems { get; private set; }

        /// <summary>
        /// GAC Listing is done
        /// </summary>
        internal bool ReadyForAction { get; private set; }

        internal string GACTitle
        {
            get
            {
                return tabPageGac.Text;
            }
            set
            {
                tabPageGac.Text = value; 
            }
        }

        internal string FileSystemTitle
        {
            get
            {
                return tabPageFileSystem.Text;
            }
            set
            {
                tabPageFileSystem.Text = value;
            }
        }

        internal string OkButtonTitle
        {
            get
            {
                return buttonOk.Text;
            }
            set 
            {
                buttonOk.Text = value;
            }
        }

        internal string CancelButtonTitle
        {
            get 
            {
                return buttonCancel.Text;
            }
            set 
            {
                buttonCancel.Text = value;
            }
        }
        #endregion

        #region Methods

        private void UpdateGACListView(GACAssembly[] gacAssemblies, bool isLastUpdate)
        { 
            if(listView1.InvokeRequired)
            {
                listView1.Invoke(new LoadAssemblyUpdateEventHandler(UpdateGACListView), new object[] { gacAssemblies, isLastUpdate });
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
                    if (!CheckGACFilter(gacAssemblies[i].Name))
                        continue;
                    _lastItemInvert = !_lastItemInvert;
                    GACAssemblyItems.Add(viewItems[i]);
                }
                listView1.Items.AddRange(viewItems);
            }

            ReadyForAction = isLastUpdate;
        }

        private void RefreshGACAssemblyFilter()
        {
            List<ListViewItem> list = new List<ListViewItem>();

            foreach (ListViewItem item in listView1.SelectedItems)
                list.Add(item);

            foreach (var item in list)
                item.Selected = false;

            listView1.Items.Clear();
            foreach (ListViewItem item in GACAssemblyItems)
            {
                if (CheckGACFilter(item.Text))
                    listView1.Items.Add(item);
            }
        }

        private bool CheckGACFilter(string name)
        {
            if (String.IsNullOrWhiteSpace(textBoxNameFilter.Text))
                return true;

            int position = name.IndexOf(textBoxNameFilter.Text, StringComparison.InvariantCultureIgnoreCase);
            return position >= 0;
        }

        #endregion

        #region Trigger

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
        
        private void textBoxNameFilter_KeyDown(object sender, KeyEventArgs e)
        {
            if (false == ReadyForAction || e.KeyCode != Keys.Return)
                return;
            RefreshGACAssemblyFilter();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPageGac)
                buttonOk.Enabled = listView1.SelectedItems.Count > 0;
            else
                buttonOk.Enabled = openFilePanel1.SelectedFiles.Length > 0;
        }

        private void openFilePanel1_SelectionChanged(object sender, FileSystemDialogs.SelectionChangedEventArgs args)
        {
            if (tabControl1.SelectedTab == tabPageFileSystem)
                buttonOk.Enabled = openFilePanel1.SelectedFiles.Length > 0; 
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
        
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            if (!ReadyForAction)
                return;

            this.DialogResult = DialogResult.OK ;
            this.Close();
        }

        #endregion
    }
}
