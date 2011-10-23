using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using Microsoft.Win32;
using System.Text;
using System.Security.Principal;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.WindowsRegistry;

namespace NetOffice.DeveloperToolbox.RegistryEditor
{
    public partial class RegistryEditorControl : UserControl, IToolboxControl
    {
        #region Fields

        UtilsRegistry _localMachine;
        UtilsRegistry _currentUser;      
        int           _currentLanguageID;
        bool          _userIsAdmin;

        #endregion

        #region Construction

        public RegistryEditorControl()
        {
            try
            {
                InitializeComponent();
                if (!DesignMode)
                {
                    _localMachine = new UtilsRegistry(Registry.LocalMachine, @"Software\Microsoft\Office");
                    _currentUser = new UtilsRegistry(Registry.CurrentUser, @"Software\Microsoft\Office");

                    _userIsAdmin = IsAdministrator();
                    pictureBoxNoAdminHint.Visible = !_userIsAdmin;
                    labelNoAdminHint.Visible = !_userIsAdmin;
                    labelNoAdminHintIcon.Visible = !_userIsAdmin;
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion

        #region Properties

        private bool IsAdministrator()
        {
            WindowsIdentity myWindowsIdentity = WindowsIdentity.GetCurrent();
            WindowsPrincipal myWindowsPrincipal = new WindowsPrincipal(myWindowsIdentity);
            return myWindowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private new bool DesignMode
        {
            get
            {
                return (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "devenv");
            }
        }

        #endregion

        #region Methods

        private void ShowKeys()
        {
            treeViewRegistry.Nodes.Clear();
            TreeNode node = treeViewRegistry.Nodes.Add("LocalMachine");
            node.Tag = _localMachine;
            foreach (UtilsRegistryKey key in _localMachine.Key.Keys)
                ShowRegNode(key, node);
          
            node = treeViewRegistry.Nodes.Add("CurrentUser");
            node.Tag = _currentUser;
            foreach (UtilsRegistryKey key in _currentUser.Key.Keys)
                ShowRegNode(key, node);
        }

        private void ShowRegNodeChilds(UtilsRegistryKey key, TreeNode node)
        {
            foreach (UtilsRegistryKey subKey in key.Keys)
                ShowRegNode(subKey,node);
        }

        private void ShowRegNode(UtilsRegistryKey key, TreeNode node)
        {           
            node = node.Nodes.Add(key.Name);
            node.Tag = true;
            if(key.Keys.Count > 0)
                node.Nodes.Add("#stub");            
        }

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

        private static TreeNode SearchChildTree(TreeView treeView, string name)
        {
            foreach (TreeNode tn in treeView.Nodes)
            {
                if (tn.Text == name)
                    return tn;
            }
            return null;
        }

        private static TreeNode SearchChildTree(TreeNode treeNode, string name)
        {
            foreach (TreeNode tn in treeNode.Nodes)
            {
                if (tn.Text == name)
                    return tn;
            }
            return null;
        }

        private UtilsRegistry GetRegistry(TreeNode node)
        {
            if (node.FullPath.StartsWith("LocalMachine", StringComparison.InvariantCultureIgnoreCase))
                return _localMachine;
            else
                return _currentUser;
        }

        private string GetFullNodePath(TreeNode node)
        {
            TreeNode rootNode = node;
            while (rootNode.Parent != null)
                rootNode = rootNode.Parent;
            UtilsRegistry registry = rootNode.Tag as UtilsRegistry;

            string path = registry.Key.Path;
            int position = node.FullPath.IndexOf("\\", StringComparison.InvariantCultureIgnoreCase);
            if (position > -1)
                path += node.FullPath.Substring(position);
            return path;
        }

        private TreeNode GetRootNode(TreeNode node)
        {
            TreeNode rootNode = node;
            while (node != null)
            {
                if (null != node.Parent)
                {
                    node = node.Parent;
                }
                else
                {
                    rootNode = node;
                    break;
                }
            }
            return rootNode;
        }

        #endregion

        #region Trigger

        private void contextMenuStripEntries_Opening(object sender, CancelEventArgs e)
        {
            try
            {
                contextMenuStripEntries.Enabled = (null != treeViewRegistry.SelectedNode);

                if (dataGridViewRegistry.SelectedCells.Count > 0)
                {
                    DataGridViewRow row = dataGridViewRegistry.Rows[dataGridViewRegistry.SelectedCells[0].RowIndex];
                    UtilsRegistryEntry entry = row.Tag as UtilsRegistryEntry;
                    if (entry == null)
                        return;
                    if (entry.Type == UtilsRegistryEntryType.Normal)
                    {
                        toolStripDeleteEntry.Enabled = true;
                        toolStripEditEntryName.Enabled = true;
                    }
                    else
                    {
                        toolStripDeleteEntry.Enabled = false;
                        toolStripEditEntryName.Enabled = false;
                    }
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void treeViewRegistry_AfterExpand(object sender, TreeViewEventArgs e)
        {
            try
            {
                e.Node.ImageIndex = 1;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void treeViewRegistry_AfterCollapse(object sender, TreeViewEventArgs e)
        {
            try
            {
                e.Node.ImageIndex = 0;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void treeViewRegistry_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            try
            {
                if ((e.Node.Nodes.Count > 0) && ("#stub" == e.Node.Nodes[0].Text))
                {
                    e.Node.Nodes.Clear();
                    string fullPath = GetFullNodePath(e.Node);
                    UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(e.Node), fullPath);
                    ShowRegNodeChilds(key, e.Node);
                }
                else if (null != e.Node.Tag)
                {
                    e.Node.Nodes.Clear();
                    string fullPath = GetFullNodePath(e.Node);
                    UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(e.Node), fullPath);
                    ShowRegNodeChilds(key, e.Node);
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void treeViewRegistry_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                TreeNode rootNode = GetRootNode(e.Node);
                UtilsRegistry regRoot = rootNode.Tag as UtilsRegistry;
                if ((regRoot.HiveKey == Registry.LocalMachine) && (!_userIsAdmin))
                    treeViewRegistry.ContextMenuStrip = contextMenuStripNoAdmin;
                else
                    treeViewRegistry.ContextMenuStrip = contextMenuStripKeys;

                toolStripKeyEdit.Enabled = (treeViewRegistry.SelectedNode != null) && (treeViewRegistry.SelectedNode.Parent != null);
                toolStripKeyDelete.Enabled = (treeViewRegistry.SelectedNode != null) && (treeViewRegistry.SelectedNode.Parent != null);

                dataGridViewRegistry.Rows.Clear();

                if (null == e.Node.Tag)
                {
                    labelCurrentPath.Text = e.Node.Text;
                    return;
                }

                string fullNodePath = GetFullNodePath(treeViewRegistry.SelectedNode);
                UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(e.Node), fullNodePath);
                bool foundDefault = false;
                foreach (UtilsRegistryEntry item in key.Entries)
                {
                    if (item.Type == UtilsRegistryEntryType.Default)
                        foundDefault = true;
                    string name = item.Name;
                    string valueType = item.ValueKind.ToString();
                    object value = item.GetValue();
                    Image typeImage = GetValueKindImage(item.ValueKind);
                    dataGridViewRegistry.Rows.Add(typeImage, name, valueType, value);
                    DataGridViewRow newRow = dataGridViewRegistry.Rows[dataGridViewRegistry.Rows.Count - 1];
                    newRow.Tag = item;
                }

                if (!foundDefault)
                {
                    UtilsRegistryEntry fakedKey = key.Entries.FakedDefaultKey;
                    string name = fakedKey.Name;
                    string valueType = fakedKey.ValueKind.ToString();
                    object value = fakedKey.GetValue();
                    Image typeImage = GetValueKindImage(fakedKey.ValueKind);
                    dataGridViewRegistry.Rows.Insert(0, typeImage, name, valueType, value);
                    DataGridViewRow newRow = dataGridViewRegistry.Rows[0];
                    newRow.Tag = fakedKey;
                }

                labelCurrentPath.Text = key.Path;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void treeViewRegistry_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            try
            {
                string fullPath = GetFullNodePath(treeViewRegistry.SelectedNode);
                UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(treeViewRegistry.SelectedNode), fullPath);
                key.Name = e.Label;
                treeViewRegistry.LabelEdit = false;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }
         
        private void buttonInfo_Click(object sender, EventArgs e)
        {
            try
            {
                InfoControl infoBox = new InfoControl("RegistryEditor.Info" + _currentLanguageID.ToString() + ".rtf", true);
                this.Controls.Add(infoBox);
                infoBox.BringToFront();
                infoBox.Show();
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                string currentPath = null;

                if (null != treeViewRegistry.SelectedNode)
                    currentPath = treeViewRegistry.SelectedNode.FullPath;

                dataGridViewRegistry.Rows.Clear();

                ShowKeys();

                if (null != currentPath)
                    RestoreExpandState(currentPath);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void toolStripKeyDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (null == treeViewRegistry.SelectedNode)
                    return;

                string fullPath = GetFullNodePath(treeViewRegistry.SelectedNode);
                TreeNode parentNode = treeViewRegistry.SelectedNode.Parent;
                if (checkBoxDeleteQuestion.Checked)
                {

                    string message = string.Format("Möchten Sie den Schlüssel <{0}> löschen?", treeViewRegistry.SelectedNode.Text);
                    DialogResult dr = MessageBox.Show(message, "Löschen bestätigen", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.No)
                        return;
                }

                UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(treeViewRegistry.SelectedNode), fullPath);
                key.Delete();
                treeViewRegistry.SelectedNode = parentNode;
                buttonRefresh_Click(this, new EventArgs());      
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void dataGridViewRegistry_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (null != treeViewRegistry.SelectedNode)
                {
                    TreeNode rootNode = GetRootNode(treeViewRegistry.SelectedNode);
                    UtilsRegistry regRoot = rootNode.Tag as UtilsRegistry;
                    if ((regRoot.HiveKey == Registry.LocalMachine) && (!_userIsAdmin))
                        dataGridViewRegistry.ContextMenuStrip = contextMenuStripNoAdmin;
                    else
                        dataGridViewRegistry.ContextMenuStrip = contextMenuStripEntries;
                }

                toolStripDeleteEntry.Enabled = (dataGridViewRegistry.SelectedCells.Count > 0);
                toolStripEditEntryName.Enabled = (dataGridViewRegistry.SelectedCells.Count > 0);
                toolStripEditEntryValue.Enabled = (dataGridViewRegistry.SelectedCells.Count > 0);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void toolStripKeyCreate_Click(object sender, EventArgs e)
        {
            try
            {
                string fullPath = GetFullNodePath(treeViewRegistry.SelectedNode);
                UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(treeViewRegistry.SelectedNode), fullPath);
                key.CreateNewSubKey();
                buttonRefresh_Click(this, new EventArgs());       
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void toolStripKeyEdit_Click(object sender, EventArgs e)
        {
            try
            {
                if (null != treeViewRegistry.SelectedNode)
                {
                    treeViewRegistry.LabelEdit = true;
                    treeViewRegistry.SelectedNode.BeginEdit();
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion

        #region DataGrid Trigger

        private void toolStripEditEntryName_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewRegistry.SelectedCells.Count == 0)
                    return;

                DataGridViewRow row = dataGridViewRegistry.Rows[dataGridViewRegistry.SelectedCells[0].RowIndex];
                UtilsRegistryEntry entry = row.Tag as UtilsRegistryEntry;

                ChangeNameDialog changeDialog = new ChangeNameDialog(entry.Name, _currentLanguageID);
                if (DialogResult.OK == changeDialog.ShowDialog(this))
                {
                    entry.Name = changeDialog.EntryNewName;
                    row.Cells[1].Value = changeDialog.EntryNewName;
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void toolStripDeleteEntry_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewRegistry.SelectedCells.Count == 0)
                    return;

                DataGridViewRow row = dataGridViewRegistry.Rows[dataGridViewRegistry.SelectedCells[0].RowIndex];

                if (checkBoxDeleteQuestion.Checked)
                {
                    string message = string.Format("Möchten Sie den Wert <{0}> löschen?", row.Cells[1].Value);
                    DialogResult dr = MessageBox.Show(message, "Löschen bestätigen", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.No)
                        return;
                }

                UtilsRegistryEntry entry = dataGridViewRegistry.Rows[dataGridViewRegistry.SelectedCells[0].RowIndex].Tag as UtilsRegistryEntry;
                entry.Delete();
          
                dataGridViewRegistry.Rows.Remove(row);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void toolStripEditEntryValue_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewRegistry.SelectedCells.Count == 0)
                    return;

                DataGridViewRow row = dataGridViewRegistry.Rows[dataGridViewRegistry.SelectedCells[0].RowIndex];
                UtilsRegistryEntry entry = row.Tag as UtilsRegistryEntry;

                switch (entry.ValueKind)
                {
                    case RegistryValueKind.ExpandString:

                    case RegistryValueKind.MultiString:
                    case RegistryValueKind.String:
                    case RegistryValueKind.Unknown:
                        ChangeStringDialog stringDialog = new ChangeStringDialog(entry.Name, entry.Value as string, _currentLanguageID);
                        if (DialogResult.OK == stringDialog.ShowDialog(this))
                        {
                            entry.Value = stringDialog.EntryValue;
                            row.Cells[3].Value = stringDialog.EntryValue;
                        }
                        break;
                    case RegistryValueKind.Binary:
                        ChangeBinaryDialog binaryDialog = new ChangeBinaryDialog(entry.Name, (entry.Value as byte[]), _currentLanguageID);
                        if (DialogResult.OK == binaryDialog.ShowDialog(this))
                        {
                            entry.Value = binaryDialog.Bytes;
                            row.Cells[3].Value = entry.GetValue();
                        }
                        break;
                    case RegistryValueKind.DWord:
                    case RegistryValueKind.QWord:
                        ChangeDWordDialog dwordDialog = new ChangeDWordDialog(entry.Name, entry.Value, _currentLanguageID);
                        if (DialogResult.OK == dwordDialog.ShowDialog(this))
                        {
                            entry.Value = dwordDialog.EntryValue;
                            row.Cells[3].Value = entry.GetValue();
                        }
                        break;

                    default:
                        throw new ArgumentException(entry.ValueKind.ToString() + " is out of range.");
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void toolStripCreateStringEntry_Click(object sender, EventArgs e)
        {
            try
            {
                string fullPath = GetFullNodePath(treeViewRegistry.SelectedNode);
                UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(treeViewRegistry.SelectedNode), fullPath);
                UtilsRegistryEntry entry = key.Keys.Add(RegistryValueKind.String, "");
                dataGridViewRegistry.Rows.Add();
                DataGridViewRow row = dataGridViewRegistry.Rows[dataGridViewRegistry.Rows.Count - 1];
                Image typeImage = GetValueKindImage(entry.ValueKind);
                row.Cells[0].Value = typeImage;
                row.Cells[1].Value = entry.Name;
                row.Cells[2].Value = entry.ValueKind.ToString();
                row.Cells[3].Value = entry.GetValue();
                row.Tag = entry;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void toolStripCreateBinaryEntry_Click(object sender, EventArgs e)
        {
            try
            {
                string fullPath = GetFullNodePath(treeViewRegistry.SelectedNode);
                UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(treeViewRegistry.SelectedNode), fullPath);
                UtilsRegistryEntry entry = key.Keys.Add(RegistryValueKind.Binary, new byte[0]);
                dataGridViewRegistry.Rows.Add();
                DataGridViewRow row = dataGridViewRegistry.Rows[dataGridViewRegistry.Rows.Count - 1];
                Image typeImage = GetValueKindImage(entry.ValueKind);
                row.Cells[0].Value = typeImage;
                row.Cells[1].Value = entry.Name;
                row.Cells[2].Value = entry.ValueKind.ToString();
                row.Cells[3].Value = entry.GetValue();
                row.Tag = entry;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void toolStripCreateDWORDEntry_Click(object sender, EventArgs e)
        {
            try
            {
                string fullPath = GetFullNodePath(treeViewRegistry.SelectedNode);
                UtilsRegistryKey key = new UtilsRegistryKey(GetRegistry(treeViewRegistry.SelectedNode), fullPath);
                UtilsRegistryEntry entry = key.Keys.Add(RegistryValueKind.DWord, 0);
                dataGridViewRegistry.Rows.Add();
                DataGridViewRow row = dataGridViewRegistry.Rows[dataGridViewRegistry.Rows.Count - 1];
                Image typeImage = GetValueKindImage(entry.ValueKind);
                row.Cells[0].Value = typeImage;
                row.Cells[1].Value = entry.Name;
                row.Cells[2].Value = entry.ValueKind.ToString();
                row.Cells[3].Value = entry.GetValue();
                row.Tag = entry;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void dataGridViewRegistry_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                TreeNode rootNode = GetRootNode(treeViewRegistry.SelectedNode);
                UtilsRegistry regRoot = rootNode.Tag as UtilsRegistry;
                if ((regRoot.HiveKey == Registry.LocalMachine) && (!_userIsAdmin))
                    return;
                toolStripEditEntryValue_Click(this, new EventArgs());
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void dataGridViewRegistry_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                TreeNode rootNode = GetRootNode(treeViewRegistry.SelectedNode);
                UtilsRegistry regRoot = rootNode.Tag as UtilsRegistry;
                if ((regRoot.HiveKey == Registry.LocalMachine) && (!_userIsAdmin))
                    return;
                if (Keys.Return == e.KeyCode)
                    toolStripEditEntryValue_Click(this, new EventArgs());
            }
            catch (Exception exception)
            {

                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion

        #region IUtilsControl Members

        public string ControlName
        {
            get { return "RegistryEditor"; }
        }
       
        public string ControlCaption
        {
            get { return "Registry Editor"; }
        }

        public System.ComponentModel.IContainer Components
        {
            get
            {
                return components;
            }
        }

        public Image Icon
        {
            get
            {
                return ReadImageFromRessource("Icon.ico");
            }
        }

        public void SetLanguage(int id)
        {
            _currentLanguageID = id;
            Translator.TranslateControls(this, "RegistryEditor.MessageTable.txt", _currentLanguageID);
        }

        public void Activate()
        {           
            buttonRefresh_Click(this, new EventArgs());
        }

        public void LoadComplete()
        {
            try
            {

            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        public void LoadConfiguration(XmlNode configNode)
        {
            try
            {
                System.Xml.XmlNode node = configNode.SelectSingleNode("AskBeforeDelete");
                if (null == node)
                {
                    node = configNode.OwnerDocument.CreateElement("AskBeforeDelete");
                    node.InnerText = "false";
                    configNode.AppendChild(node);
                }
                bool mode = Convert.ToBoolean(node.Value);
                checkBoxDeleteQuestion.Checked = mode;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        public void SaveConfiguration(XmlNode configNode)
        {
            try
            {
                System.Xml.XmlNode node = configNode.SelectSingleNode("AskBeforeDelete");
                node.InnerText = BoolToString(checkBoxDeleteQuestion.Checked);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        public new void Dispose()
        {
            base.Dispose();
        }

        #endregion       

        #region Helper Methods

        private static string BoolToString(bool b)
        {
            if (b)
                return "true";
            else
                return "false";
        }

        Image GetValueKindImage(RegistryValueKind kind)
        {
            switch (kind)
            {
                case RegistryValueKind.ExpandString:
                case RegistryValueKind.MultiString:
                case RegistryValueKind.String:
                case RegistryValueKind.Unknown:
                    return imageListValueTypes.Images[0];
                default:
                    return imageListValueTypes.Images[1];
            }
        }

        private static Image ReadImageFromRessource(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            ressourcePath = assemblyName + ".RegistryEditor." + ressourcePath;
            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            Bitmap newIcon = new Icon(ressourceStream).ToBitmap();
            return newIcon;
        }
        #endregion
    }
}
