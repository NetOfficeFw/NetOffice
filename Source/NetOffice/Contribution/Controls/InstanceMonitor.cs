using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace NetOffice.Contribution.Controls
{
    /// <summary>
    /// Realtime Instance Observer
    /// </summary>
    public partial class InstanceMonitor : UserControl
    {
        #region Fields

        private Core _factory;
        private static Color _highlightColor = Color.LightGreen;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public InstanceMonitor()
        {
            InitializeComponent();
            HighlightNodes = new Dictionary<TreeNode, DateTime>();
            View = new TreeView();
            View.Dock = DockStyle.Fill;
            Controls.Add(View);
            View.Visible = true;
            OverlayPanel = new Panel();
             
            TextBox textBox = new TextBox();
            textBox.ReadOnly = true;
            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Both;
            OverlayPanel.Controls.Add(textBox);
            textBox.Dock = DockStyle.Fill;
            textBox.Visible = true;

            Button button = new Button();
            button.Click += delegate
            {
                OverlayPanel.Visible = false;
            };
            OverlayPanel.Controls.Add(button);
            button.Location = new Point(0, 0);
            button.Width = OverlayPanel.Width;
            button.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            button.Visible = true;
            button.BringToFront();

            Controls.Add(OverlayPanel);
            OverlayPanel.Visible = false;

            HighlightTimer.Enabled = true;
            AutoExpandNodes = true;

            View.AfterSelect += delegate
            {
                if(!IsDisposed)
                    OnSelectedInstanceChanged();
            };
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs after SelectedInstance have been changed
        /// </summary>
        [Description("Occurs after SelectedInstance have been changed"), Category("InstanceMonitor")]
        public event EventHandler SelectedInstanceChanged;

        /// <summary>
        /// Occurs after Factory have been changed
        /// </summary>
        [Description("Occurs after Factory have been changed"), Category("InstanceMonitor")]
        public event EventHandler FactoryChanged;

        #endregion

        #region Properties

        /// <summary>
        /// The associated factory core
        /// </summary>
        [Description("The associated factory core"), DefaultValue(null), Category("InstanceMonitor")]
        public Core Factory
        {
            get
            {
                return _factory;
            }
            set
            {
                if (value != _factory)
                {
                    DisconnectCurrentFactory();
                    ClearView();
                    _factory = value;
                    InitializeView();
                    ConnectCurrentFactory();
                    OnFactoryChanged();
                }
            }
        }

        /// <summary>
        /// Highlight new proxies for a second
        /// </summary>
        [Description("Highlight new proxies for a second"), DefaultValue(true), Category("InstanceMonitor")]
        public bool HighlightNewProxies
        {
            get
            {
                return HighlightTimer.Enabled;
            }
            set
            {
                if (value != HighlightTimer.Enabled)
                { 
                    HighlightTimer.Enabled = value;
                    if (!value)
                    {
                        foreach (KeyValuePair<TreeNode, DateTime> item in HighlightNodes)
                            item.Key.BackColor = Color.Transparent;
                        HighlightNodes.Clear();
                    }
                }
            }
        }

        /// <summary>
        /// SelectedInstance
        /// </summary>
        [Description("Current selected instance"), DefaultValue(null), Category("InstanceMonitor")]
        public ICOMObject SelectedInstance
        {
            get
            {
                return null != View.SelectedNode ? View.SelectedNode.Tag as ICOMObject : null;
            }
        }

        /// <summary>
        /// Automaticly expand all nodes
        /// </summary>
        [Description("Automaticly expand all nodes"), DefaultValue(true), Category("InstanceMonitor")]
        public bool AutoExpandNodes { get; set; }
        
        private TreeView View { get; set; }

        private Dictionary<TreeNode, DateTime> HighlightNodes { get; set; }

        private Panel OverlayPanel { get; set; }

        private TextBox OverlayTextBox { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Raise the FactoryChanged event
        /// </summary>
        protected virtual void OnFactoryChanged()
        {
            //avoid ?.Invoke for compatibility
            if (null != FactoryChanged)
                FactoryChanged(this, EventArgs.Empty);
        }

        /// <summary>
        /// Raise the SelectedInstance event
        /// </summary>
        protected virtual void OnSelectedInstanceChanged()
        {
            //avoid ?.Invoke for compatibility
            if (null != SelectedInstanceChanged)
                SelectedInstanceChanged(this, EventArgs.Empty);
        }

        private string ComObjectName(ICOMObject comObject)
        {
            return comObject.InstanceFriendlyName;
        }

        private void ClearView()
        {
            if (InvokeRequired)
                Invoke(new Action(ClearView));
            else
                View.Nodes.Clear();
        }
        
        private void EnumerateProxies(TreeNode node, ICOMObject[] childs)
        {
            TreeNode[] childNodes = new TreeNode[childs.Length];
            for (int i = 0; i < childs.Length; i++)
            {
                ICOMObject child = childs[i];
                TreeNode childNode = new TreeNode(ComObjectName(child));
                childNodes[i] = childNode;
                EnumerateProxies(childNode, ToArray(child.ChildObjects));
            }
            if (childNodes.Length > 0)
                node.Nodes.AddRange(childNodes);
        }

        private static ICOMObject[] ToArray(IEnumerable<ICOMObject> values)
        {
            int count = 0;
            foreach (ICOMObject item in values)
                count++;
            
            ICOMObject[] result = new ICOMObject[count];
            count = 0;
            foreach (ICOMObject item in values)
            {
                result[count] = item;
                count++;
            }

            return result;
        }

        private void InitializeView()
        {
            if (InvokeRequired)
                Invoke(new Action(InitializeView));
            else
            {
                Core factory = Factory;
                if (null != factory)
                {
                    IEnumerable<ICOMObject> comObjects = factory.GetRootInstances();

                    foreach (ICOMObject comObject in comObjects)
                    {
                        TreeNode node = View.Nodes.Add(ComObjectName(comObject));
                        node.Tag = comObject.GetHashCode();
                        ICOMObject[] childs = ToArray(comObject.ChildObjects);
                        TreeNode[] childNodes = new TreeNode[childs.Length];
                        for (int i = 0; i < childs.Length; i++)
                        {
                            ICOMObject subObj = childs[i];
                            childNodes[i] = new TreeNode(ComObjectName(subObj));
                            childNodes[i].Tag = subObj.GetHashCode();
                            EnumerateProxies(childNodes[i], ToArray(subObj.ChildObjects));
                        }

                        if (childNodes.Length > 0)
                            node.Nodes.AddRange(childNodes);
                    }
                }
            }
        }

        private void DisconnectCurrentFactory()
        {
            Core factory = Factory;
            if (null != factory)
            {
                factory.ProxyAdded -= Core_ProxyAdded;
                factory.ProxyRemoved -= Core_ProxyRemoved;
                factory.ProxyCleared -= Core_ProxyCleared;
            }
        }

        private void ConnectCurrentFactory()
        {
            Core factory = Factory;
            if (null != factory)
            {
                factory.ProxyAdded += Core_ProxyAdded;
                factory.ProxyRemoved += Core_ProxyRemoved;
                factory.ProxyCleared += Core_ProxyCleared;
            }
        }

        private TreeNode FindPath(IEnumerable<ICOMObject> ownerPath)
        {
            if (null == ownerPath)
                return null;

            TreeNode targetNode = null;

            TreeNodeCollection nodes = View.Nodes;
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

        private void ShowOverlayError(Exception exception)
        {
            if (false == OverlayTextBox.Visible)
            {
                OverlayTextBox.ForeColor = Color.Red;
                OverlayTextBox.Text = Environment.NewLine + Environment.NewLine + "An unexpected error occured." + Environment.NewLine + Environment.NewLine + exception.ToString();

                OverlayPanel.Visible = true;
                OverlayPanel.BringToFront();
            }
        }

        #endregion

        #region Trigger
       
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
                        if (HighlightNewProxies)
                        {
                            newNode.BackColor = _highlightColor;
                            HighlightNodes.Add(newNode, DateTime.Now);
                        }
                        if (AutoExpandNodes)
                            View.ExpandAll();
                    }
                    else
                    {
                        TreeNode newNode = View.Nodes.Add(ComObjectName(comObject));
                        newNode.Tag = comObject.GetHashCode();
                        if (HighlightNewProxies)
                        {
                            newNode.BackColor = _highlightColor;
                            HighlightNodes.Add(newNode, DateTime.Now);
                        }
                        if (AutoExpandNodes)
                            View.ExpandAll();
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
                            if (null == item.Tag)
                                continue;
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
                    else
                    {
                        TreeNode targetNode = null;
                        foreach (TreeNode item in View.Nodes)
                        {
                            if (null == item.Tag)
                                continue;
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
                            View.Nodes.Remove(targetNode);
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
                    View.Nodes.Clear();
                    HighlightNodes.Clear();
                }
                catch (Exception exception)
                {
                    ShowOverlayError(exception);
                }
            };
            TryBeginInvoke(method);
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
