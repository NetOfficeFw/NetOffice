using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using Microsoft.Win32;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperUtils.RegistryBrowser
{
    public partial class RegistryBrowserControl : UserControl, IUtilsControl
    {
        #region Fields

        bool           _showFlag;
        KeeperRegistry _localMachine;
        KeeperRegistry _currentUser;
        InfoControl    _infoBox;

        #endregion

        #region Construction

        public RegistryBrowserControl()
        {
            InitializeComponent();
        }

        public RegistryBrowserControl(object anyTag)
        {
            InitializeComponent();
        }

        #endregion

        #region Methods

        private void SetupRegistry()
        {
            _localMachine = new KeeperRegistry(Registry.LocalMachine, @"Software\Microsoft\Office");
            _currentUser = new KeeperRegistry(Registry.CurrentUser, @"Software\Microsoft\Office");  
        }

        private void ShowKeys()
        {
            treeViewRegistry.Nodes.Clear();
            TreeNode node = treeViewRegistry.Nodes.Add("LocalMachine");
            foreach (KeeperRegistryKey key in _localMachine.Key.Keys)
                ShowRegNode(key, node);
          
            node = treeViewRegistry.Nodes.Add("CurrentUser");
            foreach (KeeperRegistryKey key in _currentUser.Key.Keys)
                ShowRegNode(key, node);
        }

        private void ShowRegNode(KeeperRegistryKey key, TreeNode node)
        {
            node = node.Nodes.Add(key.Name);
            node.Tag = key;
            foreach (KeeperRegistryKey subKey in key.Keys)
                ShowRegNode(subKey,node);
        }

        #endregion

        #region Trigger

        private void treeViewRegistry_AfterSelect(object sender, TreeViewEventArgs e)
        {
            _showFlag = true;

            dataGridViewRegistry.Rows.Clear();
            KeeperRegistryKey key = e.Node.Tag as KeeperRegistryKey;
            if (null == key)
                return;

            foreach (KeeperRegistryEntry item in key.Entries)
            {
                string name = item.Name;
                string valueType = item.ValueType.ToString();
                object value = item.Value;
                dataGridViewRegistry.Rows.Add(name, valueType, value);
                DataGridViewRow newRow = dataGridViewRegistry.Rows[dataGridViewRegistry.Rows.Count - 1];
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Tag = item;
            }

            _showFlag = false;
        }
     
        private void dataGridViewRegistry_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((_showFlag) || (e.RowIndex < 0))
                return;

            DataGridViewRow row = dataGridViewRegistry.Rows[e.RowIndex];
            KeeperRegistryEntry entry = row.Tag as KeeperRegistryEntry;
            if (null != entry)
                entry.Value = row.Cells[e.ColumnIndex].Value;
        }

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            _infoBox = new InfoControl("RegistryBrowser.Info.txt", true);
            this.Controls.Add(_infoBox);
            _infoBox.BringToFront();
            _infoBox.Show();
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            string currentPath = null;
           
            if (null != treeViewRegistry.SelectedNode)
                currentPath = treeViewRegistry.SelectedNode.FullPath;

            dataGridViewRegistry.Rows.Clear();
            SetupRegistry();
            ShowKeys();

            if (null != currentPath)
                RestoreExpandState(currentPath);
        }

        #endregion

        private void RestoreExpandState(string currentPath)
        {
            TreeNode node = null;
            string[] splitArray = currentPath.Split(new string[] { treeViewRegistry.PathSeparator }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string nodeName in splitArray)
            {
                if (node == null)
                    node = SearchChildTree(treeViewRegistry, nodeName);
                else
                    node = SearchChildTree(node, nodeName);

                if (node != null)
                    node.Expand();
            }
            treeViewRegistry.SelectedNode = node;
        }

        private TreeNode SearchChildTree(TreeView treeView, string name)
        {
            foreach (TreeNode tn in treeViewRegistry.Nodes)
            {
                if (tn.Text == name)
                    return tn;
            }
            return null;
        }

        private TreeNode SearchChildTree(TreeNode treeNode, string name)
        {
            foreach (TreeNode tn in treeNode.Nodes)
            {
                if (tn.Text == name)
                    return tn;
            }
            return null;
        } 

        #region IUtilsControl Members

        public string ControlName
        {
            get { return "RegistryBrowser"; }
        }

        public void Activate()
        {
            if (null != _infoBox)
                _infoBox.Hide();

            buttonRefresh_Click(this, new EventArgs());
        }

        public void LoadConfiguration(XmlNode configNode)
        {

        }

        public void SaveConfiguration(XmlNode configNode)
        {

        }

        public void Release()
        {

        }

        #endregion
 
    }
}
