using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using Microsoft.Win32;
using System.Text;
using System.Security.Principal;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.Utils.Registry;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
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

        #region Construction

        public RegistryEditorControl()
        {
            try
            {
                InitializeComponent();
                if (!DesignMode)
                {
                    _localMachine32 = new UtilsRegistry(Registry.LocalMachine, @"Software\Microsoft\Office");
                    _currentUser32 = new UtilsRegistry(Registry.CurrentUser, @"Software\Microsoft\Office");
                    if (Is64Bit)
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

                    _userIsAdmin = IsAdministrator();
                    _supportsInfoMessage = !_userIsAdmin;                 
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, 1033);
            }
        }

        #endregion

        #region Properties

        public bool Is64Bit
        {
            get
            {
                return System.Environment.Is64BitOperatingSystem;
            }
        }

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
                DoSearch(textBoxSearch.Text);
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

        public void SetLanguage(int id)
        {
           
        }

        public Stream GetHelpText(int lcid)
        {
            return Ressources.RessourceUtils.ReadStream("ToolboxControls.RegistryEditor.Info" + lcid.ToString() + ".rtf");
        }

        public void Activate(bool firstTime)
        {
            buttonRefresh_Click(this, new EventArgs());
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        public void Release()
        {

        }

        #endregion

        #region Methods

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

            if (Is64Bit)
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

        private void DoSearch(string expression)
        {
            UtilsRegistryKey key = null;
            UtilsRegistryKey result = null;

            if (null == treeViewRegistry.SelectedNode)
            {
                if (treeViewRegistry.Nodes.Count > 0)
                    treeViewRegistry.SelectedNode = treeViewRegistry.Nodes[0];
                else
                    return;
            }
    
            TreeNode targetNode = GetNextPossibleNode(treeViewRegistry.SelectedNode);
            string fullNodePath = GetFullNodePath(targetNode);
            key = new UtilsRegistryKey(GetRegistry(targetNode), fullNodePath);
            result = SearchHive(expression, key.Root, key);
            if (null != result)
                SelectKey(result);
            else
            {
                if (key.Root.IsLocalMachine && false == key.Root.IsWow)
                {
                    // localmachine32

                    result = SearchHive(expression, _currentUser32, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }

                    result = SearchHive(expression, _localMachine64, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }

                    result = SearchHive(expression, _currentUser64, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }
                    else
                        ShowNoResult();
                }
                else if (key.Root.IsLocalMachine && true == key.Root.IsWow)
                {
                    // localmachine64

                    result = SearchHive(expression, _currentUser64, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }

                    result = SearchHive(expression, _localMachine32, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }

                    result = SearchHive(expression, _currentUser32, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }
                    else
                        ShowNoResult();
                }
                else if (false == key.Root.IsLocalMachine && false == key.Root.IsWow)
                { 
                    // currentuser32

                    result = SearchHive(expression, _localMachine64, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }

                    result = SearchHive(expression, _currentUser64, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }

                    result = SearchHive(expression, _localMachine32, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }
                    else
                        ShowNoResult();
                }
                else
                {
                    // currentuser64

                    result = SearchHive(expression, _localMachine32, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }

                    result = SearchHive(expression, _currentUser32, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }

                    result = SearchHive(expression, _localMachine64, null);
                    if (null != result)
                    {
                        SelectKey(result);
                        return;
                    }
                    else
                        ShowNoResult();
                }
            }
        }
       
        private UtilsRegistryKey SearchHive(string expression, UtilsRegistry hive, UtilsRegistryKey selectedKeyInHive)
        {
            if (null == hive)
                return null;

            TreeNode testNode = GetRootNode(hive.IsLocalMachine, hive.IsWow);
            if(null == testNode)
                return null;

            List<UtilsRegistryKey> rootKeys = new List<UtilsRegistryKey>();
            foreach (UtilsRegistryKey item in hive.Key.Keys)
                rootKeys.Add(item);

            if ((null != selectedKeyInHive && hive.Path.Equals(selectedKeyInHive.Path, StringComparison.InvariantCultureIgnoreCase)) || null == selectedKeyInHive)
            {
                // selected key is one the 2 roots or not selected
                foreach (var item in rootKeys)
                {
                    UtilsRegistryKey resultKey = SearchTopKey(expression, item);
                    if(null != resultKey)
                        return resultKey;
                }
            }
            else
            {
                // selected key is a descendent
                int topParentIndex = 0;
                UtilsRegistryKey topParentKey = GetTopParent(selectedKeyInHive, rootKeys, out topParentIndex);

                // same top key
                UtilsRegistryKey topKeyResult = SearchTopKey(expression, topParentKey, selectedKeyInHive);
                if (null != topKeyResult)
                    return topKeyResult;

                // bottom keys
                bool ignore = true;
                foreach (var item in hive.Key.Keys)
                {
                    if (true == ignore)
                    {
                        if (item.Path == topParentKey.Path)
                            ignore = false;
                    }
                    else
                    {
                        UtilsRegistryKey bottomKeyResult = SearchTopKey(expression, item);
                        if (null != bottomKeyResult)
                            return bottomKeyResult;
                    }
                }
            }

            return null;  
        }

        private UtilsRegistryKey SearchTopKey(string expression, UtilsRegistryKey topKey)
        {
            return SearchKey(expression, topKey);
        }

        private UtilsRegistryKey SearchTopKey(string expression, UtilsRegistryKey topKey, UtilsRegistryKey selectedSubKey)
        {
            if (selectedSubKey.Path == topKey.Path)
            {
                return SearchKey(expression, topKey);
            }
            else
            {
                UtilsRegistryKey parent = CreateParentRegistryKey(selectedSubKey);
                bool ignore = true;
                foreach (var item in parent.Keys)
                {
                    if (ignore == true)
                    {
                        if (item.Path == selectedSubKey.Path)
                        { 
                            ignore = false;
                            UtilsRegistryKey res = SearchKey(expression, item);
                            if (null != res)
                                return res;
                        }
                    }
                    else
                    {
                        UtilsRegistryKey res = SearchKey(expression, item);
                        if (null != res)
                            return res;
                    }
                }

                return null;
            }
        }

        private UtilsRegistryKey SearchKey(string expression, UtilsRegistryKey key)
        {
            int position = key.Name.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase);
            if (position > -1)
                return key;

            foreach (UtilsRegistryEntry item in key.Entries)
            {
                int pos1 = item.Name.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase);
                if (pos1 > -1)
                    return key;

                string valueString = item.Value.ToString();
                int pos2 = valueString.IndexOf(expression, StringComparison.InvariantCultureIgnoreCase);
                if (pos2 > -1)
                    return key;
            }

            foreach (UtilsRegistryKey item in key.Keys)
            {
                UtilsRegistryKey resKey = SearchKey(expression, item);
                if (null != resKey)
                    return resKey;
            }

            return null;
        }

        private UtilsRegistryKey GetTopParent(UtilsRegistryKey key, List<UtilsRegistryKey> searchList, out int indexPosition)
        {
            string search = key.Path;

            int i = 0;
            foreach (var item in searchList)
            {
                string itemFull = item.Path;
                if (search.StartsWith(itemFull, StringComparison.InvariantCultureIgnoreCase))
                {
                    indexPosition = i;
                    return item;
                }
                i++;
            }

            throw new ArgumentOutOfRangeException("key");
        }

        private UtilsRegistryKey CreateParentRegistryKey(UtilsRegistryKey key)
        {
            int position = key.Path.LastIndexOf("\\");
            string path = key.Path.Substring(0, position);
            return new UtilsRegistryKey(key.Root, path);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                    object value = fakedKey.GetValue(Host.CurrentLanguageID);
                    Image typeImage = GetValueKindImage(fakedKey.ValueKind);
                    dataGridViewRegistry.Rows.Insert(0, typeImage, name, valueType, value);
                    DataGridViewRow newRow = dataGridViewRegistry.Rows[0];
                    newRow.Tag = fakedKey;
                }

                labelCurrentPath.Text = (key.Root.IsLocalMachine == true ? "[Local_Machine] " : "[Current_User] ") + key.Path;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                    if (Host.CurrentLanguageID == 1031)
                    {
                        caption = "Löschen bestätigen";
                        message = string.Format("Möchten Sie den Schlüssel löschen?{1}{1}{0}", nodeNames, Environment.NewLine);
                    }
                    else
                    {
                       caption = "Confirm";
                       message = string.Format("Delete?{1}{1}{0}", nodeNames, Environment.NewLine);
                    }

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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void textBoxSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyData != Keys.Return || String.IsNullOrWhiteSpace(textBoxSearch.Text))
                    return;
                DoSearch(textBoxSearch.Text);
               
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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

                ChangeNameDialog changeDialog = new ChangeNameDialog(entry.Name, Host.CurrentLanguageID);
                if (DialogResult.OK == changeDialog.ShowDialog(this))
                {
                    entry.Name = changeDialog.EntryNewName;
                    row.Cells[1].Value = changeDialog.EntryNewName;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                        ChangeStringDialog stringDialog = new ChangeStringDialog(entry.Name, entry.Value as string, Host.CurrentLanguageID);
                        if (DialogResult.OK == stringDialog.ShowDialog(this))
                        {
                            entry.Value = stringDialog.EntryValue;
                            row.Cells[3].Value = stringDialog.EntryValue;
                        }
                        break;
                    case RegistryValueKind.Binary:
                        ChangeBinaryDialog binaryDialog = new ChangeBinaryDialog(entry.Name, (entry.Value as byte[]), Host.CurrentLanguageID);
                        if (DialogResult.OK == binaryDialog.ShowDialog(this))
                        {
                            entry.Value = binaryDialog.Bytes;
                            row.Cells[3].Value = entry.GetValue();
                        }
                        break;
                    case RegistryValueKind.DWord:
                    case RegistryValueKind.QWord:
                        ChangeDWordDialog dwordDialog = new ChangeDWordDialog(entry.Name, entry.Value, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void RegistryEditorControl_Resize(object sender, EventArgs e)
        {
            try
            {
                labelTitle.Width = splitContainer1.Panel1.Width-24;
                checkBoxDeleteQuestion.Left = splitContainer1.Panel1.Width + 20;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception, Forms.ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
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
        
        #endregion
    }
}
