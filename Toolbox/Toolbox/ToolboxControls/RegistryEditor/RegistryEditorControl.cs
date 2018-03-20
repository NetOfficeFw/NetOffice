using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Xml;
using Microsoft.Win32;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.Utils.Registry;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    /// <summary>
    /// Registry editor clone for the ms-office hive keys
    /// </summary>
    [RessourceTable("ToolboxControls.RegistryEditor.Strings.txt")]
    public partial class RegistryEditorControl : UserControl, IToolboxControl
    {
        #region Fields

        private UtilsRegistry _localMachine32;
        private UtilsRegistry _currentUser32;
        private UtilsRegistry _localMachine64;
        private UtilsRegistry _currentUser64;
        private UtilsRegistry _localMachine;
        private UtilsRegistry _currentUser;
        private bool          _userIsAdmin;
        private bool          _supportsInfoMessage;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public RegistryEditorControl()
        {
            try
            {
                InitializeComponent();
                if (!Program.IsDesign)
                {
                    _localMachine32 = new UtilsRegistry(Registry.LocalMachine, @"Software\Microsoft\Office");
                    _currentUser32 = new UtilsRegistry(Registry.CurrentUser, @"Software\Microsoft\Office");
                    if (System.Environment.Is64BitOperatingSystem)
                    {
                        _localMachine64 = new UtilsRegistry(Registry.LocalMachine, @"Software\Wow6432Node\Microsoft\Office");
                        _currentUser64 = new UtilsRegistry(Registry.CurrentUser, @"Software\Wow6432Node\Microsoft\Office");
                        _localMachine = _localMachine64;
                        _currentUser = _currentUser64;
                    }
                    else
                    {
                        _localMachine = _localMachine32;
                        _currentUser = _currentUser32;
                    }

                    _userIsAdmin = Program.IsAdmin;
                    _supportsInfoMessage = !_userIsAdmin;

                    SearchLabel.MouseEnter += delegate
                    {
                        SearchLabel.BackColor = Color.LightGray;
                    };

                    SearchLabel.MouseLeave += delegate
                    {
                        SearchLabel.BackColor = Color.White;
                    };
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion

        #region Properties

        private string SelectedPath
        {
            get
            {
                var selNode = treeViewRegistry.SelectedNode;
                if (null != selNode)
                {
                    if(null == selNode.Parent)
                        return ((UtilsRegistry)selNode.Tag).Path;
                    else
                        return ((UtilsRegistryKey)selNode.Tag).Path;
                }
                else
                    return null;
            }
        }

        private RegistryKey SelectedRoot
        {
            get
            {
                if (null != treeViewRegistry.SelectedNode)
                {
                    TreeNode node = treeViewRegistry.SelectedNode;
                    while (null != node.Parent)
                    {
                        node = node.Parent;
                    }
                    return ((UtilsRegistry)node.Tag).HiveKey;
                }
                else
                    return null;
            }
        }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public void InitializeControl(IToolboxHost host)
        {
            Host = host;
        }

        public new void KeyDown(KeyEventArgs e)
        {
            if (false == textBoxSearch.Focused && e.KeyCode == Keys.F && e.Modifiers == Keys.Control)
                textBoxSearch.Focus();
            else if (e.KeyData == Keys.F5)
                buttonRefresh_Click(buttonRefresh, EventArgs.Empty);
            else if (e.KeyData == Keys.F3 && !String.IsNullOrWhiteSpace(textBoxSearch.Text))
            {
                if(!textBoxSearch.Focused)
                    textBoxSearch.Focus();
                DoSearchAsync(textBoxSearch.Text);
            }
        }

        public string ControlName
        {
            get { return "RegistryEditor.RegistryEditorControl"; }
        }

        public string ControlCaption
        {
            get { return "Registry"; }
        }

        public bool SupportsInfoMessage
        {
            get
            {
                return _supportsInfoMessage;
            }
        }

        public ToolboxControlMessageKind InfoMessageKind
        {
            get
            {
                return ToolboxControlMessageKind.Warning;
            }
        }

        public string InfoMessage
        {
            get
            {
                return labelNoAdminHint.Text;
            }
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
            get { return Ressources.RessourceUtils.ReadIconImageFromRessource("ToolboxControls.RegistryEditor.Icon.ico"); }
        }

        public bool SupportsHelpContent
        {
            get
            {
                return true;
            }
        }

        public Stream GetHelpText()
        {
            return Ressources.RessourceUtils.ReadStream("ToolboxControls.RegistryEditor.Info1033.rtf");
        }

        public void Activate(bool firstTime)
        {
            buttonRefresh_Click(this, new EventArgs());
            if (firstTime)
            {
                foreach (TreeNode item in treeViewRegistry.Nodes)
                    item.Expand();
            }
        }

        public void Deactivated()
        {

        }

        public void LoadComplete()
        {

        }

        public void LoadConfiguration(XmlNode configNode)
        {
            try
            {
                System.Xml.XmlNode node = configNode["AskBeforeDelete"];
                if (null == node)
                {
                    node = configNode.OwnerDocument.CreateElement("AskBeforeDelete");
                    node.InnerText = "false";
                    configNode.AppendChild(node);
                }
                bool mode = Convert.ToBoolean(node.InnerText);
                checkBoxDeleteQuestion.Checked = mode;

                node = configNode["LastSearch"];
                if (null != node)
                {
                    textBoxSearch.Text = node.InnerText;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        public void SaveConfiguration(XmlNode configNode)
        {
            try
            {
                System.Xml.XmlNode node = configNode.OwnerDocument.CreateElement("AskBeforeDelete");
                configNode.AppendChild(node);
                node.InnerText = BoolToString(checkBoxDeleteQuestion.Checked);

                if (!String.IsNullOrWhiteSpace(textBoxSearch.Text))
                {
                    node = configNode.OwnerDocument.CreateElement("LastSearch");
                    configNode.AppendChild(node);
                    node.InnerText = textBoxSearch.Text;
                }

            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        public void Release()
        {

        }

        #endregion

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {

        }

        public string NameLocalization
        {
            get
            {
                return null;
            }
        }

        #endregion

        #region Methods

        private static string BoolToString(bool b)
        {
            if (b)
                return "true";
            else
                return "false";
        }

        private Image GetValueKindImage(RegistryValueKind kind)
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

        private void ShowKeys()
        {
            treeViewRegistry.Nodes.Clear();
            TreeNode node = null;

            if (_localMachine32.Exists)
            {
                node = treeViewRegistry.Nodes.Add("LocalMachine");
                node.Tag = _localMachine32;
                foreach (UtilsRegistryKey key in _localMachine32.Key.Keys)
                    ShowRegNode(key, node);
            }

            if (_currentUser32.Exists)
            {
                node = treeViewRegistry.Nodes.Add("CurrentUser");
                node.Tag = _currentUser32;
                foreach (UtilsRegistryKey key in _currentUser32.Key.Keys)
                    ShowRegNode(key, node);
            }

            if (System.Environment.Is64BitOperatingSystem)
            {
                if (_localMachine64.Exists)
                {
                    node = treeViewRegistry.Nodes.Add("LocalMachine [Wow6432Node]");
                    node.Tag = _localMachine64;
                    foreach (UtilsRegistryKey key in _localMachine64.Key.Keys)
                        ShowRegNode(key, node);
                }

                if (_currentUser64.Exists)
                {
                    node = treeViewRegistry.Nodes.Add("CurrentUser [Wow6432Node]");
                    node.Tag = _currentUser64;
                    foreach (UtilsRegistryKey key in _currentUser64.Key.Keys)
                        ShowRegNode(key, node);
                }
            }
        }

        private void ShowRegNode(UtilsRegistryKey key, TreeNode node)
        {
            node = node.Nodes.Add(key.Name);
            node.Tag = true;
            if (key.Keys.Count > 0)
                node.Nodes.Add("#stub");
        }

        private void ShowRegNodeChilds(UtilsRegistryKey key, TreeNode node)
        {
            foreach (UtilsRegistryKey subKey in key.Keys)
            {
                TreeNode childNode = node.Nodes.Add(subKey.Name);
                childNode.Tag = subKey;
                if (subKey.Keys.Count > 0)
                    childNode.Nodes.Add("#stub");
            }
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
            TreeNode rootNode = GetRootNode(node);
            if(rootNode.Text.EndsWith("[Wow6432Node]"))
            {
                if (node.FullPath.StartsWith("LocalMachine", StringComparison.InvariantCultureIgnoreCase))
                    return _localMachine64;
                else
                    return _currentUser64;
            }
            else
            {
                if (node.FullPath.StartsWith("LocalMachine", StringComparison.InvariantCultureIgnoreCase))
                    return _localMachine32;
                else
                    return _currentUser32;
            }
        }

        private string GetNodeNames(List<TreeNode> listNodes)
        {
            string result = "";

            List<string> listNames = new List<string>();
            foreach (TreeNode item in listNodes)
                listNames.Add(item.Text);

            foreach (string item in listNames)
                result += string.Format("{0}{1}", item, Environment.NewLine);

            return result;
        }

        private string[] GetNodePaths(List<TreeNode> listNodes)
        {

            List<string> listNames = new List<string>();
            foreach (TreeNode item in listNodes)
                listNames.Add(GetFullNodePath(item));

            return listNames.ToArray();
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

        private TreeNode GetRootNode(bool machine, bool wow)
        {
            if(machine)
            {
                foreach (TreeNode item in treeViewRegistry.Nodes)
                {
                    if (true == wow && item.Text == "LocalMachine [Wow6432Node]")
                        return item;
                    else if (false == wow && item.Text == "LocalMachine")
                        return item;
                }
            }
            else
            {
                foreach (TreeNode item in treeViewRegistry.Nodes)
                {
                    if (true == wow && item.Text == "CurrentUser [Wow6432Node]")
                        return item;
                    else if (false == wow && item.Text == "CurrentUser")
                        return item;
                }
            }
            return null;
        }

        private void SelectEntryByExpression(string expression)
        {
            foreach (DataGridViewRow item in dataGridViewRegistry.Rows)
            {
                if (item.Cells[1].Value.ToString().IndexOf(expression, 0, StringComparison.InvariantCultureIgnoreCase) > -1 ||
                    item.Cells[3].Value.ToString().IndexOf(expression, 0, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    item.Selected = true;
                }
                else
                    item.Selected = false;
            }
        }

        private void SelectKey(UtilsRegistry targetKey)
        {
            if (null == targetKey)
                throw new ArgumentNullException("targetKey");
            pictureBoxNoResult.Visible = false;

            TreeNode rootNode = GetRootNode(targetKey.IsLocalMachine, targetKey.IsWow);
            UtilsRegistry rootRegistry = (UtilsRegistry)rootNode.Tag;

            if (!rootNode.IsExpanded)
                rootNode.Expand();

            TreeNode currentNode = rootNode;

            string[] array = targetKey.Path.Substring(rootRegistry.PathWithoutName.Length).Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in array)
            {
                if (!currentNode.IsExpanded)
                    currentNode.Expand();
                foreach (TreeNode node in currentNode.Nodes)
                {
                    if (node.Text == item)
                    {
                        currentNode = node;
                        continue;
                    }
                }
            }

            treeViewRegistry.SelectedNode = currentNode;
        }

        private void SelectKey(UtilsRegistryKey targetKey)
        {
            if (null == targetKey)
                throw new ArgumentNullException("targetKey");
            pictureBoxNoResult.Visible = false;

            TreeNode rootNode = GetRootNode(targetKey.Root.IsLocalMachine, targetKey.Root.IsWow);

            if (!rootNode.IsExpanded)
                rootNode.Expand();

            TreeNode currentNode = rootNode;

            string[] array = targetKey.Path.Substring(targetKey.Root.Path.Length+1).Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in array)
            {
                if (!currentNode.IsExpanded)
                    currentNode.Expand();
                foreach (TreeNode node in currentNode.Nodes)
                {
                    if (node.Text == item)
                    {
                        currentNode = node;
                        continue;
                    }
                }
            }

            treeViewRegistry.SelectedNode = currentNode;
        }

        private void ShowNoResult()
        {
            pictureBoxNoResult.Visible = true;
        }

        private TreeNode GetNextPossibleNode(TreeNode node)
        {
            if (node.Tag is UtilsRegistry)
                return node;

            if (null != node.FirstNode)
            {
                if (node.FirstNode.Text.EndsWith("#stub"))
                {
                    node.Expand();
                    node.Collapse();
                }
                if (null != node.FirstNode)
                    return node.FirstNode;
            }

            if (null != node.NextNode)
                return node.NextNode;

            TreeNode currentNode = node;

            while (null != currentNode.Parent)
            {
                if (node.Tag is UtilsRegistry)
                    return null;
                if (null != currentNode.Parent.NextNode)
                    return currentNode.Parent.NextNode;
                currentNode = currentNode.Parent;
            }

            return null;
        }

        //private UtilsRegistryKey FirstRootAsKey()
        //{
        //    var firstRoot = FirstRootKey();
        //    return new UtilsRegistryKey(null, firstRoot.Path);
        //}

        //private UtilsRegistry FirstRootKey()
        //{
        //    if (treeViewRegistry.Nodes.Count > 0)
        //        return (UtilsRegistry)treeViewRegistry.Nodes[0].Tag;
        //    else
        //        return null;
        //}

        private string FirstPath()
        {
            if (treeViewRegistry.Nodes.Count > 0)
                return ((UtilsRegistry)treeViewRegistry.Nodes[0].Tag).Path;
            else
                return null;
        }

        private RegistryKey FirstRoot()
        {
            if (treeViewRegistry.Nodes.Count > 0)
                return( (UtilsRegistry)treeViewRegistry.Nodes[0].Tag).HiveKey;
            else
                return null;
        }

        private IEnumerable<UtilsRegistry> AvailableRootKeys()
        {
            List<UtilsRegistry> result = new List<UtilsRegistry>();

            foreach (TreeNode item in treeViewRegistry.Nodes)
                result.Add((UtilsRegistry)item.Tag);

            return result;
        }

        private bool IsResultSelected(string expression)
        {
            if (null == treeViewRegistry.SelectedNode)
                return false;

            foreach (DataGridViewRow row in dataGridViewRegistry.Rows)
            {
                if (dataGridViewRegistry.SelectedRows.Contains(row))
                {
                    string name = row.Cells[1].Value.ToString();
                    string value = row.Cells[3].Value.ToString();

                    bool matchName = name.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase) > -1;
                    bool matchValue = value.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase) > -1;

                    if (matchName || matchValue)
                        return true;
                }
                else
                {
                    string name = row.Cells[1].Value.ToString();
                    string value = row.Cells[3].Value.ToString();

                    bool matchName = name.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase) > -1;
                    bool matchValue = value.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase) > -1;

                    if (matchName || matchValue)
                        return false;
                }
            }

            UtilsRegistry registry = treeViewRegistry.SelectedNode.Tag as UtilsRegistry;
            if (null != registry)
            {
                bool nameMatch = registry.Name.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase) > -1;
                if (nameMatch)
                    return true;
            }
            else
            {
                UtilsRegistryKey key = treeViewRegistry.SelectedNode.Tag as UtilsRegistryKey;
                bool nameMatch = key.Name.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase) > -1;
                if (nameMatch)
                    return true;
            }

            return false;
        }

        private bool Locked { get; set; }
        private void LockUI()
        {
            treeViewRegistry.Enabled = false;
            dataGridViewRegistry.Enabled = false;
            Locked = true;
        }

        private void UnlockUI()
        {
            treeViewRegistry.Enabled = true;
            dataGridViewRegistry.Enabled = true;
            Locked = false;
        }

        private RegistrySearch Search { get; set; }

        private void DoSearchAsyncComplete(IAsyncResult result)
        {
            try
            {
                Action method = result.AsyncState as Action;
                method.EndInvoke(result);
                Action invokeUI = delegate
                {
                    pictureBoxSearching.Visible = false;
                    UnlockUI();
                    if (null != Search.Result)
                    {
                        SelectKey(Search.Result);
                        SelectEntryByExpression(Search.Expression);
                    }
                    else
                        ShowNoResult();
                };
                Invoke(invokeUI);
            }
            catch
            {
                ;
            }
        }

        private void DoSearchAsync(string expression)
        {
            if (Locked)
                return;

            var rootKeys = AvailableRootKeys();
            bool isResultSelected = IsResultSelected(expression);
            RegistryKey startFromRootKey = null != SelectedRoot ? SelectedRoot : FirstRoot();
            string startFromPath = null != SelectedPath ? SelectedPath : FirstPath();
            bool isTopLevelPath = AvailableRootKeys().Any(e=>e.Path == startFromPath);

            Search = new RegistrySearch(rootKeys, startFromRootKey, null != SelectedPath ? SelectedPath : FirstPath(), isTopLevelPath, expression, isResultSelected);

            Action method = delegate
            {
                Search.Search();
            };
            pictureBoxNoResult.Visible = false;
            pictureBoxSearching.Visible = true;
            LockUI();
            method.BeginInvoke(DoSearchAsyncComplete, method);
        }

        #endregion

        #region Trigger

        private void treeViewRegistry_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyData == Keys.Delete)
                    toolStripKeyDelete_Click(sender, new EventArgs());
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                toolStripKeyExport.Enabled = (treeViewRegistry.SelectedNode != null) && (treeViewRegistry.SelectedNode.Parent != null);

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

                labelCurrentPath.Text = (key.Root.IsLocalMachine == true ? "[Local_Machine] " : "[Current_User] ") + key.Path;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                string currentPath = null;
                string[] entries = new string[dataGridViewRegistry.SelectedRows.Count];
                for (int i = 0; i < dataGridViewRegistry.SelectedRows.Count; i++)
                    entries[i] = dataGridViewRegistry.SelectedRows[i].Cells[1].Value.ToString();
                if (null != treeViewRegistry.SelectedNode)
                    currentPath = treeViewRegistry.SelectedNode.FullPath;

                dataGridViewRegistry.Rows.Clear();

                ShowKeys();

                if (null != currentPath)
                    RestoreExpandState(currentPath);

                foreach (DataGridViewRow item in dataGridViewRegistry.Rows)
                {
                    if (entries.Contains(item.Cells[1].Value.ToString()))
                        item.Selected = true;
                    else
                        item.Selected = false;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void toolStripKeyDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (0 == treeViewRegistry.SelectedNodes.Count)
                    return;

                string[] fullPathSelectedNodes = GetNodePaths(treeViewRegistry.SelectedNodes);

                TreeNode parentNode = null;
                if( (null != treeViewRegistry.SelectedNode.PrevNode) && (treeViewRegistry.SelectedNodes.Count == 1))
                    parentNode = treeViewRegistry.SelectedNode.PrevNode;
                else
                    parentNode = treeViewRegistry.SelectedNode.Parent;

                if (checkBoxDeleteQuestion.Checked)
                {
                    string nodeNames = GetNodeNames(treeViewRegistry.SelectedNodes);
                    string message = null;
                    string caption = null;
                    caption = "Confirm";
                    message = string.Format("Delete?{1}{1}{0}", nodeNames, Environment.NewLine);

                    DialogResult dr = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.No)
                        return;
                }

                UtilsRegistry registryRoot = GetRegistry(treeViewRegistry.SelectedNode);
                foreach (string fullPath in fullPathSelectedNodes)
                {
                    UtilsRegistryKey key = new UtilsRegistryKey(registryRoot, fullPath);
                    key.Delete();
                }

                treeViewRegistry.SelectedNode = parentNode;
                buttonRefresh_Click(this, new EventArgs());
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void toolStripKeyExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (null != treeViewRegistry.SelectedNode && (!(treeViewRegistry.SelectedNode.Tag is UtilsRegistry)))
                {
                    SaveFileDialog dlg = new SaveFileDialog();
                    dlg.Filter = "*.reg|*.reg";
                    if (DialogResult.OK == dlg.ShowDialog(this))
                    {
                        string fullPath = GetFullNodePath(treeViewRegistry.SelectedNode);
                        UtilsRegistry reg = GetRegistry(treeViewRegistry.SelectedNode);
                        UtilsRegistryKey key = new UtilsRegistryKey(reg, fullPath);
                        UtilsRegistryKeyExporter.Export(dlg.FileName, reg.InnerKey.ToString() + "\\" + key.Path);
                    }
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void textBoxSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyData != Keys.Return || String.IsNullOrWhiteSpace(textBoxSearch.Text))
                    return;
                DoSearchAsync(textBoxSearch.Text);
                //DoSearch(textBoxSearch.Text);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }


        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                pictureBoxNoResult.Visible = false;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void textBoxSearch_Leave(object sender, EventArgs e)
        {
            try
            {
                 pictureBoxNoResult.Visible = false;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void RegistryEditorControl_Resize(object sender, EventArgs e)
        {
            try
            {
                checkBoxDeleteQuestion.Left = splitContainer1.Panel1.Width + 20;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void SearchLabel_Click(object sender, EventArgs e)
        {
            try
            {
                DoSearchAsync(textBoxSearch.Text);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
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

                ChangeNameDialog changeDialog = new ChangeNameDialog(entry.Name, 1033);
                if (DialogResult.OK == changeDialog.ShowDialog(this))
                {
                    entry.Name = changeDialog.EntryNewName;
                    row.Cells[1].Value = changeDialog.EntryNewName;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
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
                    string message = string.Format("Do you want delete the value <{0}> ?", row.Cells[1].Value);
                    DialogResult dr = MessageBox.Show(message, "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.No)
                        return;
                }

                UtilsRegistryEntry entry = dataGridViewRegistry.Rows[dataGridViewRegistry.SelectedCells[0].RowIndex].Tag as UtilsRegistryEntry;
                entry.Delete();

                dataGridViewRegistry.Rows.Remove(row);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                        ChangeStringDialog stringDialog = new ChangeStringDialog(entry.Name, entry.Value as string, 1033);
                        if (DialogResult.OK == stringDialog.ShowDialog(this))
                        {
                            entry.Value = stringDialog.EntryValue;
                            row.Cells[3].Value = stringDialog.EntryValue;
                        }
                        break;
                    case RegistryValueKind.Binary:
                        ChangeBinaryDialog binaryDialog = new ChangeBinaryDialog(entry.Name, (entry.Value as byte[]), 1033);
                        if (DialogResult.OK == binaryDialog.ShowDialog(this))
                        {
                            entry.Value = binaryDialog.Bytes;
                            row.Cells[3].Value = entry.GetValue();
                        }
                        break;
                    case RegistryValueKind.DWord:
                    case RegistryValueKind.QWord:
                        ChangeDWordDialog dwordDialog = new ChangeDWordDialog(entry.Name, entry.Value, 1033);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
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
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void dataGridViewRegistry_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                TreeNode rootNode = GetRootNode(treeViewRegistry.SelectedNode);
                if (null == rootNode)
                    return;
                UtilsRegistry regRoot = rootNode.Tag as UtilsRegistry;
                if ((regRoot.HiveKey == Registry.LocalMachine) && (!_userIsAdmin))
                    return;
                switch (e.KeyCode)
                {
                    case Keys.Return:
                        toolStripEditEntryValue_Click(this, new EventArgs());
                        break;
                    case Keys.Delete:
                        toolStripDeleteEntry_Click(this, new EventArgs());
                        break;
                    default:
                        break;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}