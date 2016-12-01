using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using NetOffice.Tools;
using NetOffice.Misc;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Shows various diagnostics and monitor proxy management
    /// </summary>
    internal partial class TrayMenuMonitorItemControl : UserControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="comAddin">owner addin</param>
        public TrayMenuMonitorItemControl(COMAddinBase comAddin)
        {
            if (null == comAddin)
                throw new ArgumentNullException("comAddin");
            InitializeComponent();

            Addin = comAddin;
            ViewOptions = new TrayMenuMonitorItem.Options(EnumeratorGrid, HierarchicalGrid, HighlightTimer, OnDisableHighlight);
            AutoExpandCheckBox.DataBindings.Add("Checked", ViewOptions, "AutoExpandNodes", false, DataSourceUpdateMode.OnPropertyChanged);
            HighlightCheckBox.DataBindings.Add("Checked", ViewOptions, "HighlightNewNodes", false, DataSourceUpdateMode.OnPropertyChanged);     
            ShowTimeColumnCheckBox.DataBindings.Add("Checked", ViewOptions, "ShowConsoleTimeColumn", false, DataSourceUpdateMode.OnPropertyChanged);
            ShowKindColumnCheckBox.DataBindings.Add("Checked", ViewOptions, "ShowConsoleKindColumn", false, DataSourceUpdateMode.OnPropertyChanged);
            
            HighlightNodes = new Dictionary<TreeNode, DateTime>();

            ContentPanel.Controls.Remove(OverlayPanel);
            Controls.Add(OverlayPanel);
            OverlayPanel.Dock = DockStyle.Fill;

            CoreRadioButton.Tag = TrayMenuMonitorItem.ShowMode.Core;
            SettingsRadioButton.Tag = TrayMenuMonitorItem.ShowMode.Settings;
            ConsoleRadioButton.Tag = TrayMenuMonitorItem.ShowMode.Console;
            DiagnosticsRadioButton.Tag = TrayMenuMonitorItem.ShowMode.Diagnostics;
            ProxiesRadioButton.Tag = TrayMenuMonitorItem.ShowMode.Proxies;
            OptionsRadioButton.Tag = TrayMenuMonitorItem.ShowMode.Options;

            Messages = new SortableBindingList<DebugConsole.ConsoleMessage>(Addin.Factory.Console.MessagesInternal);
            ShownItems = new Dictionary<RadioButton, object>();
            ShownItems.Add(CoreRadioButton, Addin.Factory);
            ShownItems.Add(SettingsRadioButton, Addin.Factory.Settings);
            ShownItems.Add(ConsoleRadioButton, Messages);
            ShownItems.Add(DiagnosticsRadioButton, new NetOffice.Misc.SelfDiagnostics(comAddin));
            ShownItems.Add(ProxiesRadioButton, comAddin.Roots);
            ShownItems.Add(OptionsRadioButton, ViewOptions);

            EnumeratorGrid.Dock = DockStyle.Fill;
            SingleGrid.Dock = DockStyle.Fill;
            HierarchicalGrid.Dock = DockStyle.Fill;
            OptionsGrid.Dock = DockStyle.Fill;

            HighlightTimer.Enabled = true;

            HeaderRadioButton_CheckedChanged(CoreRadioButton, EventArgs.Empty);
        }

        #endregion
        
        #region Properties

        /// <summary>
        /// View Options
        /// </summary>
        public TrayMenuMonitorItem.Options ViewOptions { get; set; }

        private TrayMenuMonitorItem.ShowMode Mode { get; set; }

        private RadioButton SelectedHeaderButton
        {
            get
            {
                foreach (KeyValuePair<RadioButton, object> item in ShownItems)
                {
                    if (item.Key.Checked)
                        return item.Key;
                }
                return null;
            }
        }

        private SortableBindingList<DebugConsole.ConsoleMessage> Messages { get; set; }

        private Dictionary<RadioButton, object> ShownItems { get; set; }

        private COMAddinBase Addin { get; set; }

        private Dictionary<TreeNode, DateTime> HighlightNodes { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Set current view mode
        /// </summary>
        /// <param name="mode">target mode</param>
        public void SetCurrentMode(TrayMenuMonitorItem.ShowMode mode)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<TrayMenuMonitorItem.ShowMode>(SetCurrentMode), mode);
            }
            else
            {
                RadioButton button = GetHeaderButton(mode);
                if(!button.Visible)
                    HeaderRadioButton_CheckedChanged(button, EventArgs.Empty);
            }
        }

        /// <summary>
        /// Get current view mode
        /// </summary>
        /// <returns>view mode</returns>
        public TrayMenuMonitorItem.ShowMode GetCurrentMode()
        {
            if (InvokeRequired)
            {
                return (TrayMenuMonitorItem.ShowMode)Invoke(new Func<TrayMenuMonitorItem.ShowMode>(GetCurrentMode));
            }
            else
            {
                return (TrayMenuMonitorItem.ShowMode)SelectedHeaderButton.Tag;
            }
        }

        /// <summary>
        /// Set View Category Visibility
        /// </summary>
        /// <param name="mode">target mode, Core is unsupported</param>
        /// <param name="visible">visibilty</param>
        public void SetModeVisible(TrayMenuMonitorItem.ShowMode mode, bool visible)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<TrayMenuMonitorItem.ShowMode, bool>(SetModeVisible), mode, visible);
            }
            else
            {
                if (mode != TrayMenuMonitorItem.ShowMode.Core)
                {
                    RadioButton button = GetHeaderButton(mode);
                    if (button == SelectedHeaderButton)
                        HeaderRadioButton_CheckedChanged(CoreRadioButton, EventArgs.Empty);
                    button.Visible = visible;
                }
            }
        }

        /// <summary>
        /// Get Category Visibility
        /// </summary>
        /// <param name="mode">target mode</param>
        /// <returns>visibilty</returns>
        public bool GetModeVisible(TrayMenuMonitorItem.ShowMode mode)
        {
            if (InvokeRequired)
            {
                return (bool)Invoke(new Func<TrayMenuMonitorItem.ShowMode, bool>(GetModeVisible), mode);
            }
            else
            {
                RadioButton button = GetHeaderButton(mode);
                return button.Visible;
            }
        }

        private RadioButton GetHeaderButton(TrayMenuMonitorItem.ShowMode mode)
        {
            foreach (KeyValuePair<RadioButton, object> item in ShownItems)
            {
                TrayMenuMonitorItem.ShowMode itemMode = (TrayMenuMonitorItem.ShowMode)item.Key.Tag;
                if (itemMode == mode)
                    return item.Key;
            }
            throw new ArgumentOutOfRangeException();
        }

        private void UpdateMode()
        {
            TrayMenuMonitorItem.ShowMode mode = (TrayMenuMonitorItem.ShowMode)SelectedHeaderButton.Tag;
            Mode = mode;
        }

        private void ConnectCurrent()
        {
            switch (Mode)
            {
                case TrayMenuMonitorItem.ShowMode.Core:
                    Addin.Factory.ProxyCountChanged += Core_ProxyCountChanged;
                    Addin.Factory.IsInitializedChanged += Core_IsInitializedChanged;
                    break;
                case TrayMenuMonitorItem.ShowMode.Settings:
                    break;
                case TrayMenuMonitorItem.ShowMode.Console:
                    Addin.Factory.Console.MessageAdded += Console_MessageAdded;
                    Addin.Factory.Console.MessageRemoved += Console_MessageRemoved;
                    Addin.Factory.Console.MessageClear += Console_MessageClear;
                    break;
                case TrayMenuMonitorItem.ShowMode.Diagnostics:
                    break;
                case TrayMenuMonitorItem.ShowMode.Proxies:
                    Addin.Factory.ProxyAdded += Core_ProxyAdded;
                    Addin.Factory.ProxyRemoved += Core_ProxyRemoved;
                    Addin.Factory.ProxyCleared += Core_ProxyCleared;
                    break;
                case TrayMenuMonitorItem.ShowMode.Options:
                    break;
                default:
                    throw new IndexOutOfRangeException();
            }
        }

        private void Disconnect(TrayMenuMonitorItem.ShowMode mode)
        {
            switch (mode)
            {
                case TrayMenuMonitorItem.ShowMode.Core:
                    Addin.Factory.ProxyCountChanged -= Core_ProxyCountChanged;
                    Addin.Factory.IsInitializedChanged -= Core_IsInitializedChanged;
                    break;
                case TrayMenuMonitorItem.ShowMode.Settings:
                    break;
                case TrayMenuMonitorItem.ShowMode.Console:
                    Addin.Factory.Console.MessageAdded -= Console_MessageAdded;
                    Addin.Factory.Console.MessageRemoved -= Console_MessageRemoved;
                    Addin.Factory.Console.MessageClear -= Console_MessageClear;
                    break;
                case TrayMenuMonitorItem.ShowMode.Diagnostics:
                    break;
                case TrayMenuMonitorItem.ShowMode.Proxies:
                    Addin.Factory.ProxyAdded -= Core_ProxyAdded;
                    Addin.Factory.ProxyRemoved -= Core_ProxyRemoved;
                    Addin.Factory.ProxyCleared -= Core_ProxyCleared;
                    break;
                case TrayMenuMonitorItem.ShowMode.Options:
                    break;
                default:
                    throw new IndexOutOfRangeException();
            }
        }

        private void EnumerateProxies(TreeNode node, ICOMObject[] childs)
        {
            TreeNode[] childNodes = new TreeNode[childs.Length];
            for (int i = 0; i < childs.Length; i++)
            {
                ICOMObject child = childs[i];
                TreeNode childNode = new TreeNode(ComObjectName(child));
                childNodes[i] = childNode;
                EnumerateProxies(childNode, child.ChildObjects.ToArray());
            }
            if (childNodes.Length > 0)
                node.Nodes.AddRange(childNodes);
        }

        private void ShowDataSource(object dataSource)
        {
            if (dataSource is TrayMenuMonitorItem.Options)
            {
                EnumeratorGrid.Visible = false;
                SingleGrid.Visible = false;
                HierarchicalGrid.Visible = false;
                OptionsGrid.Visible = true;
            }
            else if (dataSource is IEnumerable<ICOMObject>)
            {
                EnumeratorGrid.Visible = false;
                SingleGrid.Visible = false;
                HierarchicalGrid.Visible = true;
                OptionsGrid.Visible = false;

                IEnumerable<ICOMObject> comObjects = dataSource as IEnumerable<ICOMObject>;
                HierarchicalGrid.Nodes.Clear();

                foreach (ICOMObject comObject in comObjects)
                {
                    TreeNode node = HierarchicalGrid.Nodes.Add(ComObjectName(comObject));
                    node.Tag = comObject.GetHashCode();
                    ICOMObject[] childs = comObject.ChildObjects.ToArray();
                    TreeNode[] childNodes = new TreeNode[childs.Length];
                    for (int i = 0; i < childs.Length; i++)
                    {
                        ICOMObject subObj = childs[i];
                        childNodes[i] = new TreeNode(ComObjectName(subObj));
                        childNodes[i].Tag = subObj.GetHashCode();
                        EnumerateProxies(childNodes[i], subObj.ChildObjects.ToArray());
                    }

                    if (childNodes.Length > 0)
                        node.Nodes.AddRange(childNodes);
                }

                if (ViewOptions.AutoExpandNodes)
                    HierarchicalGrid.ExpandAll();
            }
            else if (dataSource is IEnumerable)
            {
                EnumeratorGrid.Visible = true;
                SingleGrid.Visible = false;
                HierarchicalGrid.Visible = false;
                OptionsGrid.Visible = false;
                EnumeratorGrid.DataSource = dataSource;

                switch (EnumeratorGrid.Columns.Count)
                {
                    case 1:
                        EnumeratorGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                        EnumeratorGrid.Columns[0].Width = EnumeratorGrid.Width - 32;                       
                        break;
                    case 2:
                        EnumeratorGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                        EnumeratorGrid.Columns[0].Width = (EnumeratorGrid.Width / 2) - 16;
                        EnumeratorGrid.Columns[1].Width = (EnumeratorGrid.Width / 2) - 16;                      
                        break;
                    case 3:
                        EnumeratorGrid.Columns[0].Width = FirstColumnOptimumWidht();
                        EnumeratorGrid.Columns[1].Visible = ViewOptions.ShowConsoleTimeColumn;
                        EnumeratorGrid.Columns[2].Visible = ViewOptions.ShowConsoleKindColumn;
                        if (!EnumeratorGrid.Columns[1].Visible || !EnumeratorGrid.Columns[2].Visible)
                            EnumeratorGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        else
                            EnumeratorGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                        break;
                    case 0:
                        break;
                    default:
                        break;
                }
            }
            else
            {
                SingleGrid.SelectedObject = dataSource;
                EnumeratorGrid.Visible = false;
                SingleGrid.Visible = true;
                HierarchicalGrid.Visible = false;
                OptionsGrid.Visible = false;
            }
        }

        private int FirstColumnOptimumWidht()
        {
            int result = (EnumeratorGrid.Width - Columns2Widht()) - 32;
            if (result < 0)
                result = 20;
            return result;
        }

        private int Columns2Widht()
        {
            int result = 0;
            for (int i = 1; i < EnumeratorGrid.Columns.Count; i++)
                result += EnumeratorGrid.Columns[i].Width;
            return result;
        }

        private void ShowEmpty()
        {
            EnumeratorGrid.Visible = false;
            EnumeratorGrid.Visible = false;
            SingleGrid.Visible = false;
            HierarchicalGrid.Visible = false;
            OptionsGrid.Visible = false;
        }

        private string ComObjectName(ICOMObject comObject)
        {
            return comObject.InstanceFriendlyName;
        }

        private TreeNode FindPath(IEnumerable<ICOMObject> ownerPath)
        {
            if (null == ownerPath)
                return null;

            TreeNode targetNode = null;

            TreeNodeCollection nodes = HierarchicalGrid.Nodes;
            foreach (ICOMObject comObject in ownerPath)
            {
                if (null == nodes)
                    return null;
                targetNode = null;

                int comObjectHashCode = comObject.GetHashCode();
                foreach (TreeNode node in nodes)
                {
                    int nodeHashCode = (int)node.Tag;
                    if (comObjectHashCode == nodeHashCode)
                    {
                        nodes = node.Nodes;
                        targetNode = node;
                        break;
                    }
                }

                if (null == targetNode)
                    return null;
            }

            return targetNode;
        }

        private void OnDisableHighlight()
        {
            foreach (KeyValuePair<TreeNode, DateTime> item in HighlightNodes)
            {
                item.Key.BackColor = Color.Transparent;
            }
            HighlightNodes.Clear();
        }

        private void ShowOverlayError(Exception exception)
        {
            if (false == OverlayTextBox.Visible)
            {
                OverlayTextBox.ForeColor = Color.Red;
                OverlayTextBox.Text = "An unexpected error occured." + Environment.NewLine + Environment.NewLine + exception.ToString();
                OverlayPanel.Visible = true;
                OverlayPanel.BringToFront();

            }
        }

        private void ShowOverlayText(string text)
        {
            if (false == OverlayTextBox.Visible)
            {
                OverlayTextBox.ForeColor = Color.FromKnownColor(KnownColor.ControlText);
                OverlayTextBox.Text = text;
                OverlayPanel.Visible = true;
                OverlayPanel.BringToFront();

            }
        }

        private void TryBeginInvoke(Action method)
        {
            try
            {
                if (null != Parent && Parent.IsHandleCreated)
                    BeginInvoke(method);
            }
            catch
            {
                ;
            }
        }

        #endregion
        
        #region Trigger

        private void HeaderRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                RadioButton button = sender as RadioButton;
                if (button.Checked)
                {
                    Disconnect(Mode);
                    object dataSource = null;
                    if (ShownItems.TryGetValue(button, out dataSource))
                    {
                        UpdateMode();
                        ConnectCurrent();
                        ShowDataSource(dataSource);
                    }
                    else
                        ShowEmpty();
                }
                else
                    ShowEmpty();
            }
            catch (Exception exception)
            {
                ShowOverlayError(exception);
            }       
        }

        private void Core_ProxyAdded(Core sender, IEnumerable<ICOMObject> ownerPath, ICOMObject comObject)
        {
            Action method = delegate
            {
                try
                {
                    TreeNode node = FindPath(ownerPath);
                    if (null != node)
                    {
                        TreeNode newNode = node.Nodes.Add(ComObjectName(comObject));
                        newNode.Tag = comObject.GetHashCode();
                        if (ViewOptions.HighlightNewNodes)
                        {
                            newNode.BackColor = Color.LightGreen;
                            HighlightNodes.Add(newNode, DateTime.Now);
                        }
                        if (ViewOptions.AutoExpandNodes)
                            HierarchicalGrid.ExpandAll();
                    }
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
        }

        private void Core_ProxyRemoved(Core sender, IEnumerable<ICOMObject> ownerPath, ICOMObject comObject)
        {
            Action method = delegate
            {
                try
                {
                    int objectHashCode = comObject.GetHashCode();
                    TreeNode node = FindPath(ownerPath);
                    if (null != node)
                    {
                        TreeNode targetNode = null;
                        foreach (TreeNode item in node.Nodes)
                        {
                            int itemHashCode = (int)item.Tag;
                            if (itemHashCode == objectHashCode)
                            {
                                targetNode = item;
                                break;
                            }
                        }
                        if (null != targetNode)
                        {
                            HighlightNodes.Remove(targetNode);
                            node.Nodes.Remove(targetNode);
                        }
                    }
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
        }

        private void Core_ProxyCleared(Core sender)
        {
            Action method = delegate
            {
                try
                {
                    HierarchicalGrid.Nodes.Clear();
                    HighlightNodes.Clear();
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
        }

        private void Console_MessageAdded(DebugConsole sender, DebugConsole.ConsoleMessage message)
        {
            Action method = delegate
            {
                try
                {
                    Messages.Add(message);
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
        }

        private void Console_MessageRemoved(DebugConsole sender, DebugConsole.ConsoleMessage message, int index)
        {
            Action method = delegate
            {
                try
                {
                    Messages.Remove(message);
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
        }

        private void Console_MessageClear(DebugConsole sender)
        {
            Action method = delegate
            {
                try
                {
                    Messages.Clear();
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
        }

        private void Core_IsInitializedChanged(bool isInitialized)
        {
            Action method = delegate
            {
                try
                {
                    if (SelectedHeaderButton == CoreRadioButton)
                        SingleGrid.Refresh();
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
        }

        private void Core_ProxyCountChanged(int proxyCount)
        {
            Action method = delegate
            {
                try
                {
                    if (SelectedHeaderButton == CoreRadioButton)
                        SingleGrid.Refresh();
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
        }

        private void EnumeratorGrid_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (EnumeratorGrid.SelectedRows.Count == 0)
                    return;

                string text = String.Empty;

                ITypedList list = EnumeratorGrid.DataSource as ITypedList;
                if (null != list)
                {
                    PropertyDescriptorCollection properties = list.GetItemProperties(null);
                    if (null != properties)
                    {
                        object item = EnumeratorGrid.SelectedRows[0].DataBoundItem;
                        Type itemType = item.GetType();

                        foreach (PropertyDescriptor descriptor in properties)
                        {
                            object value = itemType.InvokeMember(descriptor.Name,
                                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance, null, item, null);
                            if (null != value)
                                text += descriptor.Name + ":" + value.ToString() + Environment.NewLine;
                        }
                    }
                }

                if (text != String.Empty)
                    ShowOverlayText(text);
            }
            catch (Exception exception)
            {
                ShowOverlayError(exception);
            }         
        }

        private void CloseOverlayButton_Click(object sender, EventArgs e)
        {
            try
            {
                OverlayPanel.Visible = false;
            }
            catch (Exception exception)
            {
                ShowOverlayError(exception);
            }
        }

        private void HighlightTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                DateTime now = DateTime.Now;             
                foreach (KeyValuePair<TreeNode, DateTime> item in HighlightNodes)
                {
                    if ((now - item.Value).TotalMilliseconds >= 1000)
                    {
                        item.Key.BackColor = Color.Transparent;
                    }
                }
            }
            catch (Exception exception)
            {
                HighlightTimer.Enabled = false;
                ShowOverlayError(exception);
            }
        }

        #endregion    
    }
}
