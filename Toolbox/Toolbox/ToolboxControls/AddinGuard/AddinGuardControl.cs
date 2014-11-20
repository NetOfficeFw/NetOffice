using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using Microsoft.Win32;
using System.Text;
using System.Security.Principal;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.AddinGuard
{
    [RessourceTable("ToolboxControls.AddinGuard.Strings.txt")]
    public partial class AddinGuardControl : UserControl, IToolboxControl
    {
        #region Fields

        private AddinsKey       _addinsItemToDisplay;
        private DisabledKey     _disabledItemToDisplay;
        private Exception       _displayedException;
        private WatchController _controller;
        private bool            _programmaticChange;
        private string          _message;
        private bool            _boolFirstFiredMessage = true;

        #endregion

        #region Construction

        public AddinGuardControl()
        {
            try
            {
                InitializeComponent();
                if (!DesignMode)
                {
                    panelDeactivatedElements.Location = panelRegistryValues.Location;
                    panelDeactivatedElements.Size = panelRegistryValues.Size;

                    labelColorLegendCaption.Location = labelIconLegendCaption.Location;
                    splitContainer1.Panel2.Controls.Add(panelInfos);
                    panelInfos.Location = panelRegistryValues.Location;
                    panelInfos.Size = panelRegistryValues.Size;
                    panelInfos.Visible = true;

                    panelColorLegend.Location = panelIconLegend.Location;
                    panelColorLegend.Size = panelIconLegend.Size;

                    _controller = new WatchController();
                    _controller.PropertyChanged += new PropertyChangedEventHandler(_controller_PropertyChanged);
                    _controller.WatchNotify.MessageFired += new EventHandler(WatchNotify_MessageFired);

                    pictureBoxNoAdmin.Visible = !IsAdministrator();
                    labelNoAdminHint.Visible = !IsAdministrator();
                    labelNoAdminHintIcon.Visible = !IsAdministrator();
                    _controller.ReadOnlyModeForMachineKeys = !IsAdministrator();
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, 1033);
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

        private TreeNode SelectedWatcherNode
        {
            get
            {
                if (null == treeViewRegistry.SelectedNode)
                    return null;

                AddinsKey item = treeViewRegistry.SelectedNode.Tag as AddinsKey;
                if (null != item)
                    return treeViewRegistry.SelectedNode;

                item = treeViewRegistry.SelectedNode.Parent.Tag as AddinsKey;
                if (null != item)
                    return treeViewRegistry.SelectedNode.Parent;

                return null;
            }
        }

        private AddinsKey CurrentDisplayedItem
        {
            get
            {
                if (null == treeViewRegistry.SelectedNode)
                    return null;

                AddinsKey item = treeViewRegistry.SelectedNode.Tag as AddinsKey;
                if (null != item)
                    return item;

                return treeViewRegistry.SelectedNode.Parent.Tag as AddinsKey;
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
        }

        public string ControlName
        {
            get { return "AddinGuard.AddinGuardControl"; }
        }
    
        public string ControlCaption
        {
            get { return "Addin Guard"; }
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
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.AddinGuard.Icon.png"); }
        }

        public bool SupportsHelpContent
        {
            get
            {
                return true;
            }
        }

        public void Activate(bool firstTime)
        {

        }

        public void Deactivated()
        {

        }
        
        public void LoadComplete()
        {
 
        }

        public bool SupportsInfoMessage
        {
            get
            {
                return false;
            }
        }

        public ToolboxControlMessageKind InfoMessageKind
        {
            get
            {
                return ToolboxControlMessageKind.Uncategorized;
            }
        }

        public string InfoMessage
        {
            get
            {
                return String.Empty;
            }
        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
            if (configNode.ChildNodes.Count == 0)
                configNode.InnerXml = Ressources.RessourceUtils.ReadString("ToolboxControls.AddinGuard.DefaultConfiguration.txt");


            System.Xml.XmlNode node = configNode.SelectSingleNode("Active");
            bool mode = Convert.ToBoolean(node.Value);
            if (mode)
                radioButtonActivate.Checked = true;

            node = configNode.SelectSingleNode("SetLoadBehavior");
            mode = Convert.ToBoolean(node.Value);
            if (mode)
                checkBoxRestoreLoadBehavior.Checked = true;


            node = configNode.SelectSingleNode("TrayNotify");
            mode = Convert.ToBoolean(node.Value);
            if (mode)
                radioButtonTray.Checked = true;
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
            if (configNode.ChildNodes.Count == 0)
                configNode.InnerXml = Ressources.RessourceUtils.ReadString("ToolboxControls.AddinGuard.DefaultConfiguration.txt");

            System.Xml.XmlNode node = configNode.SelectSingleNode("Active");
            node.InnerText  = BoolToString(radioButtonActivate.Checked);

            node = configNode.SelectSingleNode("SetLoadBehavior");
            node.InnerText = BoolToString(checkBoxRestoreLoadBehavior.Checked);

            node = configNode.SelectSingleNode("TrayNotify");
            node.InnerText = BoolToString(radioButtonTray.Checked);
        }

        public void SetLanguage(int id)
        {
            //_controller.ActiveLanguageID = id;
            //Translation.Translator.TranslateControls(this, "ToolboxControls.AddinGuard.MessageTable.txt", id);
        }

        public Stream GetHelpText(int lcid)
        {
            Translation.ToolLanguage language = Host.Languages[lcid, false];
            if(null != language)
            {
                string content = language.Components["AddinGuard-Help"].ControlRessources["richTextBoxHelpContent"].Value2;
                return Ressources.RessourceUtils.CreateStreamFromString(content);
            }
            else
                return Ressources.RessourceUtils.ReadStream("ToolboxControls.AddinGuard.Info" + lcid.ToString() + ".rtf");
        }

        public void Release()
        {
            _controller.Dispose();
        }

        #endregion

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {

        }

        public void Localize(Translation.ItemCollection strings)
        {
            Translation.Translator.TranslateControls(this, strings);
        }

        public void Localize(string name, string text)
        {
            Translation.Translator.TranslateControl(this, name, text);
        }

        public string GetCurrentText(string name)
        {
            return Translation.Translator.TryGetControlText(this, name);
        }

        public string NameLocalization
        {
            get
            {
                return null;
            }
        }

        public IEnumerable<ILocalizationChildInfo> Childs
        {
            get
            {
                return new ILocalizationChildInfo[]{new LocalizationDefaultChildInfo("Help", typeof(Controls.InfoLayer.InfoControl))};
            }
        }

        #endregion

        #region PropertyChanged Trigger

        private void WatchNotify_MessageFiredInvoke()
        {
            try
            {
                if (_boolFirstFiredMessage)
                {
                    labelMessages.Clear();
                    _boolFirstFiredMessage = false;
                }
                labelMessages.Text = _message.Replace("\r\n", " ") + Environment.NewLine + labelMessages.Text;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void WatchNotify_MessageFired(object sender, EventArgs e)
        {
            try
            {
                _message = sender as string;
                this.Invoke(new MethodInvoker(WatchNotify_MessageFiredInvoke));
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void WatchNotify_ExceptionThrownInvoke()
        {
            try
            {
                if (_boolFirstFiredMessage)
                {
                    labelMessages.Clear();
                    _boolFirstFiredMessage = false;
                }
                labelMessages.Text = "Exception:" + _displayedException.Message + Environment.NewLine + labelMessages.Text;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void _controller_AddinsKeyChangedInvoke()
        {
            try
            {
                string selectedKey = null;
                if (null != treeViewRegistry.SelectedNode)
                    selectedKey = treeViewRegistry.SelectedNode.Name;

                if (null == _addinsItemToDisplay)
                    return;

                TreeNode node = treeViewRegistry.Nodes[_addinsItemToDisplay.Name];
                DeleteAddinNodes(node, _addinsItemToDisplay.RootKey);
                
                foreach (AddinKey subItem in _addinsItemToDisplay.Addins)
                {
                    string key = subItem.Parent.RootKey.ToString() + "-" + subItem.Parent.Name + "-" + subItem.Name;
                    TreeNode subNode = node.Nodes.Add(key, subItem.Name);
                    if (subItem.Parent.RootKey == Registry.LocalMachine)
                        subNode.ImageIndex = 2;
                    else
                        subNode.ImageIndex = 3;
                    subNode.SelectedImageIndex = subNode.ImageIndex;

                    subNode.BackColor = ToColor(subItem.LoadBehavior);

                    subNode.Tag = subItem;
                }

                treeViewRegistry.ExpandAll();

                if (null != selectedKey)
                    SelectNode(treeViewRegistry, selectedKey);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }
 
        private void _controller_DisabledKeyChangedInvoke()
        {
            try
            {
                string selectedKey = null;
                if (null != treeViewRegistry.SelectedNode)
                    selectedKey = treeViewRegistry.SelectedNode.Name;

                if (null == _disabledItemToDisplay)
                    return;

                TreeNode node = treeViewRegistry.Nodes[_disabledItemToDisplay.Name];
                DeleteDisabledNodes(node);
                foreach (DisabledValue subItem in _disabledItemToDisplay.Values)
                {
                    string key = subItem.Parent.RootKey.ToString() + "-" + subItem.Parent.Name + "-" + subItem.Name;
                    TreeNode subNode = node.Nodes.Add(subItem.Name);
                    subNode.ImageIndex = 4;
                    subNode.SelectedImageIndex = subNode.ImageIndex;
                    subNode.Tag = subItem;
                }

                treeViewRegistry.ExpandAll();

                if (null != selectedKey)
                    SelectNode(treeViewRegistry, selectedKey);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void _controller_PropertyChangedInvoke()
        {
            try
            {
                _programmaticChange = true;
                radioButtonActivate.Checked = _controller.Enabled;
                checkBoxRestoreLoadBehavior.Checked = _controller.RestoreLastLoadBehavior;
                radioButtonMsgBox.Checked = (NotificationType.MessageBox == _controller.NotifyType);
                _programmaticChange = false;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void _controller_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            try
            {
                _addinsItemToDisplay = sender as AddinsKey;
                _disabledItemToDisplay = sender as DisabledKey;
                _displayedException = sender as Exception;

                if (null != _addinsItemToDisplay)
                    this.Invoke(new MethodInvoker(_controller_AddinsKeyChangedInvoke));
                else if (null != _disabledItemToDisplay)
                    this.Invoke(new MethodInvoker(_controller_DisabledKeyChangedInvoke));
                else if (sender is WatchController)
                    this.Invoke(new MethodInvoker(_controller_PropertyChangedInvoke));
                else if (sender is Exception)
                    this.Invoke(new MethodInvoker(WatchNotify_ExceptionThrownInvoke));
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }  

        #endregion

        #region Gui Trigger

        private void checkBoxRestoreLoadBehavior_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (_programmaticChange)
                    return;

                CheckBox button = sender as CheckBox;
                _controller.RestoreLastLoadBehavior = button.Checked;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void radioButtonMsgBox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (_programmaticChange)
                    return;

                RadioButton button = sender as RadioButton;
                if (button.Checked)
                    _controller.NotifyType = NotificationType.MessageBox;
                else
                    _controller.NotifyType = NotificationType.TrayBallon;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void radioButtonActivate_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                RadioButton button = sender as RadioButton;

                if (button.Checked)
                    checkBoxRestoreLoadBehavior.ForeColor = Color.Blue;
                else
                    checkBoxRestoreLoadBehavior.ForeColor = Color.Gray;

                if (_programmaticChange)
                    return;

                _controller.Enabled = button.Checked;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void treeViewRegistry_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                dataGridViewValues.Rows.Clear();
                if (e.Node.Tag is AddinKey)
                {
                    panelDeactivatedElements.Visible = false;
                    panelInfos.Visible = false;
                    panelRegistryValues.Visible = true;

                    AddinKey addinKey = e.Node.Tag as AddinKey;
                    foreach (AddinKeyValue item in addinKey.Values)
                    {
                        dataGridViewValues.Rows.Add();
                        DataGridViewRow row = dataGridViewValues.Rows[dataGridViewValues.Rows.Count - 1];
                        row.Cells[0].Value = GetValueKindImage(item.Type);
                        row.Cells[1].Value = item.Name;
                        row.Cells[2].Value = item.Type;
                        row.Cells[3].Value = item.Value;
                    }
                }
                else if (e.Node.Tag is DisabledValue)
                {
                    panelRegistryValues.Visible = false;
                    panelInfos.Visible = false;
                    panelDeactivatedElements.Visible = true;
                    DisabledValue disabledValue = e.Node.Tag as DisabledValue;

                    labelOfficeProduct.Text = disabledValue.OfficeProductVersion;
                    labelDisabledRegistryValue.Text = disabledValue.Value;
                    labelDisabledRegistryPath.Text = disabledValue.Parent.RegistryPath + " - " + disabledValue.ValueName;
                }
                else
                {
                    panelRegistryValues.Visible = false;
                    panelDeactivatedElements.Visible = false;
                    panelInfos.Visible = true;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
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
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void buttonChangeLegend_Click(object sender, EventArgs e)
        {
            try
            {
                Button button = sender as Button;
                if (">" == button.Text)
                {
                    panelIconLegend.Visible = false;
                    panelColorLegend.Visible = true;
                    labelColorLegendCaption.Visible = true;
                    labelIconLegendCaption.Visible = false;
                    button.Text = "<";
                }
                else
                {
                    panelIconLegend.Visible = true;
                    panelColorLegend.Visible = false;
                    labelColorLegendCaption.Visible = false;
                    labelIconLegendCaption.Visible = true;
                    button.Text = ">";
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        #endregion

        #region Methods

        private Image GetValueKindImage(RegistryValueKind kind)
        {
            switch (kind)
            {
                case RegistryValueKind.ExpandString:
                case RegistryValueKind.MultiString:
                case RegistryValueKind.String:
                case RegistryValueKind.Unknown:
                    return imageListEntries.Images[0];
                default:
                    return imageListEntries.Images[1];
            }
        }

        private static Color ToColor(int? loadBehavior)
        {
            if (null == loadBehavior)
                return Color.Red;

            if (loadBehavior == 2)
                return Color.Yellow;

            if (loadBehavior == 0)
                return Color.Orange;

            if ((loadBehavior != 0) && (loadBehavior != 1) && (loadBehavior != 2) && (loadBehavior != 3) && (loadBehavior != 8) && (loadBehavior != 9) && (loadBehavior != 16))
                return Color.DarkKhaki;

            return Color.Transparent;
        }


        private static void SelectNode(TreeView treeView, string key)
        {
            foreach (TreeNode node in treeView.Nodes)
            {
                if (key == node.Name)
                {
                    treeView.SelectedNode = node;
                    return;
                }
                foreach (TreeNode subNode in node.Nodes)
                {
                    if (key == subNode.Name)
                    {
                        treeView.SelectedNode = subNode;
                        return;
                    }
                }
            }
        }

        private void DeleteDisabledNodes(TreeNode node)
        {
            List<TreeNode> deleteList = new List<TreeNode>();
            foreach (TreeNode childNode in node.Nodes)
            {
                if (childNode.Tag is DisabledValue)
                    deleteList.Add(childNode);
            }

            foreach (TreeNode childNode in deleteList)
                node.Nodes.Remove(childNode);
        }

        private void DeleteAddinNodes(TreeNode node, RegistryKey rootKey)
        {
            List<TreeNode> deleteList = new List<TreeNode>();
            foreach (TreeNode childNode in node.Nodes)
            {
                if (childNode.Tag is AddinKey)
                {
                    AddinKey nodeKey = childNode.Tag as AddinKey;
                    if(nodeKey.Parent.RootKey == rootKey)
                        deleteList.Add(childNode);
                }
            }

            foreach (TreeNode childNode in deleteList)
                node.Nodes.Remove(childNode);
        }

        private void UpdateSubKeys(TreeNode node)
        {
            AddinsKey item = node.Tag as AddinsKey;
            if(null != item)
            {
                RegistryKey key = item.RootKey.OpenSubKey(item.RegistryPath);
                if (null != key)
                {
                    node.Nodes.Clear();
                    string[] subKeyNames = key.GetSubKeyNames();
                    foreach (string subKeyName in subKeyNames)
                    {
                        TreeNode subNode = node.Nodes.Add(subKeyName);
                        subNode.ForeColor = Color.Gray;
                    }
                    key.Close();
                }
            }
        }
        
        private void SetControlsEnabled(bool enabled)
        {
            foreach (Control control in splitContainer1.Panel2.Controls)
            {
                control.Enabled = enabled;
                if (!enabled)
                {
                    control.Text = "";
                    control.BackColor = Color.Gray;
                }

                foreach (Control subControl in control.Controls)
                {
                    subControl.Enabled = enabled;
                    if (!enabled)
                    {
                        subControl.Text = "";
                        subControl.BackColor = Color.Gray;
                    }
                }
            }
        }
       
        #endregion

        #region Static Methods

        private static string BoolToString(bool b)
        {
            if (b)
                return "true";
            else
                return "false";
        }

        #endregion
    }
}
