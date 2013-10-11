using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOTools.FileSystemDialogs
{
    public partial class OpenFilePanel : UserControl
    { 
        public OpenFilePanel()
        {
            InitializeComponent();

            NodeDesktop = TreeView1.Nodes[0];
            NodeMyComputer = TreeView1.Nodes[1];
            NodeMyDocuments = TreeView1.Nodes[2];
            NodeSpecialFolders = TreeView1.Nodes[3];
            NodeTemplateFolders = TreeView1.Nodes[4];

            Misc = new MiscSettings(MiscSettings_PropertyChanged);
            Default = new DefaultSettings(Misc, DefaultableSettings_PropertyChanged);
            Desktop = new DesktopSettings(Default, DefaultableSettings_PropertyChanged);
            MyComputer = new MyComputerSettings(Default, DefaultableSettings_PropertyChanged);
            MyDocuments = new MyDocumentsSettings(Default, DefaultableSettings_PropertyChanged);
            SpecialFolders = new SpecialFoldersSettings(Default, DefaultableSettings_PropertyChanged);
            TemplateFolders = new TemplateFoldersSettings(Default, DefaultableSettings_PropertyChanged);

            FolderTemplates = new FolderTemplateCollection();
            FileSystemHandler = new FileSystemManager();
          
            ToolStripContainer1_ContentPanel_SizeChanged(this, new EventArgs());
            ShowAll();
        }
         
        #region Properties

        [Category("Settings"), DisplayName("Misc"), Description("Provides misc settings."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public MiscSettings Misc { get; private set; }

        [Category("Settings"), DisplayName("Default"), Description("Provides default settings."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public DefaultSettings Default { get; private set; }

        [Category("Settings"), DisplayName("Desktop"), Description("Provides desktop settings."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public DesktopSettings Desktop { get; private set; }

        [Category("Settings"), DisplayName("MyComputer"), Description("Provides machine settings."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public MyComputerSettings MyComputer { get; private set; }

        [Category("Settings"), DisplayName("MyDocuments"), Description("Provides my documents settings."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public MyDocumentsSettings MyDocuments { get; private set; }

        [Category("Settings"), DisplayName("SpecialFolders"), Description("Provides my special folders settings."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public SpecialFoldersSettings SpecialFolders { get; private set; }

        [Category("Settings"), DisplayName("TemplateFolders"), Description("Provides my template folder settings."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public TemplateFoldersSettings TemplateFolders { get; private set; }

        private TreeNode NodeDesktop { get; set; }

        private TreeNode NodeMyComputer { get; set; }

        private TreeNode NodeMyDocuments { get; set; }

        private TreeNode NodeSpecialFolders { get; set; }

        private TreeNode NodeTemplateFolders { get; set; }

        public FolderTemplateCollection FolderTemplates { get; private set; }

        internal FileSystemManager FileSystemHandler { get; private set; }

        #endregion

        #region Show Methods

        internal void ClearView()
        {
            NodeDesktop.Nodes.Clear();
            NodeMyComputer.Nodes.Clear();
            NodeMyDocuments.Nodes.Clear();
            NodeSpecialFolders.Nodes.Clear();
            NodeTemplateFolders.Nodes.Clear();
            ListView1.Items.Clear();
        }

        internal void ShowAll()
        {
            ClearView();
            ShowDesktop();
            ShowMyComputer();
            ShowMyDocuments();
            ShowSpecialFolders();
            ShowTemplateFolders();
            ApplyViewSettings();
        }

        internal void ShowDesktop()
        {
            FileSystemInfo fsi = FileSystemHandler.GetDesktop();
            NodeDesktop.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeDesktop);
            if (settings.HasAllowedSubFolders(fsi))
            { 
                NodeDesktop.Nodes.Add("?:");
                NodeDesktop.Collapse();
            }
        }

        internal void ShowMyComputer()
        {
            FileSystemInfo fsi = FileSystemHandler.GetMyComputer();
            NodeMyComputer.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeMyComputer);
            if (settings.HasAllowedSubFolders(fsi))
            { 
                NodeMyComputer.Nodes.Add("?:");
                NodeMyComputer.Collapse();
            }
        }

        internal void ShowMyDocuments()
        {
            FileSystemInfo fsi = FileSystemHandler.GetMyDocuments();
            NodeMyDocuments.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeMyDocuments);
            if (settings.HasAllowedSubFolders(fsi))
            { 
                NodeMyDocuments.Nodes.Add("?:");
                NodeMyDocuments.Collapse();
            }
        }

        internal void ShowSpecialFolders()
        {
            FileSystemInfo fsi = FileSystemHandler.GetSpecialFolders();
            NodeSpecialFolders.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeSpecialFolders);
            if (settings.HasAllowedSubFolders(fsi))
            { 
                NodeSpecialFolders.Nodes.Add("?:");
                NodeSpecialFolders.Collapse();
            }
        }

        internal void ShowTemplateFolders()
        {
          
        }

        #endregion

        #region ApplySettings Methods

        private void ApplyMiscSettings(string propertyName = null)
        {
            comboBoxFileTypes.DataSource = Misc.Filters;
            splitContainer1.Panel1Collapsed = !Misc.ShowCategoryPanel;
            switch (Misc.SelectedCategory)
            {
 // TODO: Sichtbarkeit prüfen
                case RootCategory.Desktop:
                    TreeView1.SelectedNode = TreeView1.Nodes[0];
                    break;
                case RootCategory.MyComputer:
                    TreeView1.SelectedNode = TreeView1.Nodes[1];
                    break;
                case RootCategory.MyDocuments:
                    TreeView1.SelectedNode = TreeView1.Nodes[2];
                    break;
                case RootCategory.SpecialFolders:
                    TreeView1.SelectedNode = TreeView1.Nodes[3];
                    break;
                case RootCategory.TemplateFolders:
                    TreeView1.SelectedNode = TreeView1.Nodes[4];
                    break;
                default:
                    break;
            }
            TreeView1_AfterSelect(TreeView1, new TreeViewEventArgs(TreeView1.SelectedNode));
        }

        private void UdpateRootCategoryVisibility(TreeNode node, DefaultableSettings settings, int nodeIndex)
        {
            if (settings.GetRuntimeValue("Visible"))
            {
                if (!TreeView1.Nodes.Contains(node))
                    TreeView1.Nodes.Insert(nodeIndex, node);
            }
            else
            {
                if (TreeView1.Nodes.Contains(node))
                    TreeView1.Nodes.Remove(node);
            }
        }

        private void ApplyDefaultableSettings(string propertyName = null)
        {
// TODO: Indexe dynamisch
            UdpateRootCategoryVisibility(NodeDesktop, Desktop, 0);
            UdpateRootCategoryVisibility(NodeMyComputer, MyComputer, 1);
            UdpateRootCategoryVisibility(NodeMyDocuments, MyDocuments, 2);
            UdpateRootCategoryVisibility(NodeSpecialFolders, SpecialFolders, 3);
            UdpateRootCategoryVisibility(NodeTemplateFolders, TemplateFolders, 4);
            ShowAll();
        }
    
        #endregion

        #region ApplyViewSettings Methods

        private void ApplyViewSettings()
        {
            DefaultableSettings settings = GetCurrentSettings();
            if (null != settings && settings.GetRuntimeValue("AllowBrowseFolders"))
            {
                StripButtonCreateDirectory.Visible = settings.GetRuntimeValue("AllowAddFolders");
                StripButtonDeleteDirectory.Visible = settings.GetRuntimeValue("AllowDeleteFolders");
                StripButtonDeleteFile.Visible = settings.GetRuntimeValue("AllowDeleteFiles");
            }
            else
            {
                StripButtonCreateDirectory.Visible = false;
                StripButtonDeleteDirectory.Visible = false;
                StripButtonDeleteFile.Visible = false;
            }
        }

        #endregion

        private RootCategory GetCurrentSelectedRootCategory(TreeNode node = null)
        {
            TreeNode selectedNode = node;
            if(null == selectedNode)
                selectedNode = TreeView1.SelectedNode;
            if (null == selectedNode)
                return RootCategory.Undefined;

            while (selectedNode != null)
            {
                if (selectedNode.Parent == null)
                {
                    switch (selectedNode.Name)
                    {
                        case "NodeDesktop":
                            return RootCategory.Desktop;
                        case "NodeMyComputer":
                            return RootCategory.MyComputer;
                        case "NodeMyDocuments":
                            return RootCategory.MyDocuments;
                        case "NodeSpecialFolders":
                            return RootCategory.SpecialFolders;
                        case "NodeTemplateFolders":
                            return RootCategory.TemplateFolders;
                        default:
                            break;
                    }
                }
                selectedNode = selectedNode.Parent;
            }

            return RootCategory.Undefined;
        }

        private int CalculateDriveImageIndex(DrvInfo item)
        {
            switch (item.Type)
            {
                case System.IO.DriveType.CDRom:
                    return item.IsReady ? 2 : 3;
                case System.IO.DriveType.Network:
                    return item.IsReady ? 4 : 5;
                default:
                    return item.IsReady ? 0 : 1;
            }
        }

        private void GotoSubNode(TreeNode node, string subNodeName)
        {
            if (!node.IsExpanded)
                node.Expand();

            foreach (TreeNode item in node.Nodes)
            {
                if (item.Text == subNodeName)
                {
                    TreeView1.SelectedNode = item;
                    return;
                }
            }
        }

        private void ExpandNode(TreeNode node, FileSystemInfo fsInfo)
        {
            node.Nodes.Clear();
            DefaultableSettings settings = GetCurrentSettings(node);
            if (!settings.HasAllowedSubFolders(fsInfo))
                return;

            foreach (var item in fsInfo.Drives)
            {
                if (settings.AllowShowDrive(item))
                { 
                    TreeNode subNode = node.Nodes.Add(item.Name);
                    subNode.Tag = item;
                    subNode.ImageIndex = 12;

                    if (GetCurrentSettings(subNode).HasAllowedSubFolders(item))
                        subNode.Nodes.Add("?:");
                }
            }

            foreach (var item in fsInfo.Directories)
            {
                if (settings.AllowShowFolder(item))
                { 
                    TreeNode subNode = node.Nodes.Add(item.Name);
                    subNode.Tag = item;
                    subNode.ImageIndex = 12;

                    if (GetCurrentSettings(subNode).HasAllowedSubFolders(item))
                        subNode.Nodes.Add("?:");
                }
            }
        }

        private void ShowCurrentNodeInListView(TreeNode node, FileSystemInfo fsInfo)
        {
            DefaultableSettings settings = GetCurrentSettings(node);
            if (settings.GetRuntimeValue("AllowBrowseFolders"))
            { 
                foreach (DrvInfo item in fsInfo.Drives)
                {
                    if (settings.AllowShowDrive(item))
                    {
                        ListViewItem viewItem = ListView1.Items.Add(item.Name);
                        viewItem.Tag = item;
                        viewItem.ImageIndex = 1;
                    }
                }

                foreach (FolderInfo item in fsInfo.Directories)
                {
                    if (settings.AllowShowFolder(item))
                    { 
                        ListViewItem viewItem = ListView1.Items.Add(item.Name);
                        viewItem.Tag = item;
                        viewItem.ImageIndex = 1;
                    }
                }
            }

            foreach (FiInfo item in fsInfo.Files)
            {
                if (settings.AllowShowFile(item))
                { 
                    ListViewItem viewItem = ListView1.Items.Add(item.Name);
                    viewItem.SubItems.Add(item.ValidatedSize);
                    viewItem.Tag = item;
                    viewItem.ImageIndex = 0;
                }
            }

            StripButtonCreateDirectory.Enabled = settings.CanCreateFolders(fsInfo);
            StripButtonDeleteDirectory.Enabled = false;
            StripButtonDeleteFile.Enabled = false;
        }

        private DefaultableSettings GetCurrentSettings(TreeNode node = null)
        { 
            RootCategory category = GetCurrentSelectedRootCategory(node);
            switch (category)
            {
                case RootCategory.Desktop:
                    return Desktop;
                case RootCategory.MyComputer:
                    return MyComputer;
                case RootCategory.MyDocuments:
                    return MyDocuments;
                case RootCategory.SpecialFolders:
                    return SpecialFolders;
                case RootCategory.TemplateFolders:
                    return TemplateFolders;
                default:
                    return null;
            }
        }

        private void UpdateToolStripButtons()
        {
            if (ListView1.SelectedItems.Count > 0)
            {
                foreach (ListViewItem item in ListView1.SelectedItems)
                {
                    if (item.Tag is FolderInfo)
                        StripButtonDeleteDirectory.Enabled = true;
                    else if (item.Tag is FiInfo)
                        StripButtonDeleteFile.Enabled = true;
                }
            }
            else
            {
                StripButtonDeleteDirectory.Enabled = false;
                StripButtonDeleteFile.Enabled = false;
            }
        }

        #region Trigger
        
        private void MiscSettings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            ApplyMiscSettings(e.PropertyName);
        }
       
        private void DefaultableSettings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            ApplyDefaultableSettings(e.PropertyName);
        }

        private void ToolStripContainer1_ContentPanel_SizeChanged(object sender, EventArgs e)
        {
            splitContainer1.Size = new Size(splitContainer1.Parent.Width - 2, splitContainer1.Parent.Height - 68);
        }

        private void TreeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            FileSystemInfo fsInfo = e.Node.Tag as FileSystemInfo;
            if (null == fsInfo)
                return;

            DefaultableSettings settings = GetCurrentSettings(e.Node);
            if (!settings.HasAllowedSubFolders(fsInfo))
            {
                e.Node.Nodes.Clear();
                return;
            }

            ExpandNode(e.Node, fsInfo);
        }

        private void TreeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {           
            if (null == e.Node)
                return;
            ListView1.Items.Clear();
            FileSystemInfo fsInfo = e.Node.Tag as FileSystemInfo;
            if (null == fsInfo)
                return;
            
            ShowCurrentNodeInListView(e.Node, fsInfo);
            ApplyViewSettings();
        }

        private void ListView1_DoubleClick(object sender, EventArgs e)
        {
            if (ListView1.SelectedItems.Count == 0)
                return;

            FileSystemInfo fsInfo = ListView1.SelectedItems[0].Tag as FileSystemInfo;
            if (fsInfo is FiInfo)
            {
            }
            else
            {
                GotoSubNode(TreeView1.SelectedNode, fsInfo.Name);
            }
        }
         
        private void ListView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            UpdateToolStripButtons();
        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateToolStripButtons();
        }

        private void comboBoxFileTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            Misc.CurrentFilter = comboBoxFileTypes.SelectedItem as FileFilterItem;
            if (null != TreeView1.SelectedNode)
            {
                ListView1.Items.Clear();
                ShowCurrentNodeInListView(TreeView1.SelectedNode, TreeView1.SelectedNode.Tag as FileSystemInfo);
            }
        }

        private void StripButtonViewLargeIcon_Click(object sender, EventArgs e)
        {
            ListView1.View = View.LargeIcon;
        }

        private void StripButtonViewSmallIcon_Click(object sender, EventArgs e)
        {
            ListView1.View = View.Tile;
        }

        private void StripButtonViewDetails_Click(object sender, EventArgs e)
        {
            ListView1.View = View.Details;
        }

        private void OpenFilePanel_Resize(object sender, EventArgs e)
        {
            colHeaderSize.Width = 100;
            colHeaderName.Width = ListView1.Width - (colHeaderSize.Width+10);
        }

        #endregion
    }
}
