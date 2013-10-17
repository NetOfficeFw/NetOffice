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
    /// <summary>
    /// Custom OpenFile dialog as usercontrol
    /// </summary>
    [Designer(typeof(OpenFilePanelDesigner)), ToolboxBitmap(typeof(OpenFileDialog)), Description("Custom OpenFile dialog as usercontrol.")]
    public partial class OpenFilePanel : UserControl
    {
        #region Ctor

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
            Misc.SetCategoryPanelWidth(splitContainer1.Panel1.Width);
            Localization = new OpenFileLocalization(LocalizationSettings_PropertyChanged);
            Localization.RaisePropertyChanged("");
            FileSystemHandler = new FileSystemManager();

            ToolStripContainer1_ContentPanel_SizeChanged(this, new EventArgs());
            ShowAll();
            this.Load += new EventHandler(OpenFilePanel_Load);
        }

        #endregion

        #region Events

        [Category("!OpenFile"), Description("Occurs when a file is (doubleclicked by user.")]
        public event FileDoubleClickEventHandler FileDoubleClick;

        private void RaiseFileDoubleClick(string file)
        {
            if(null != FileDoubleClick)
            FileDoubleClick(this, new FileDoubleClickEventArgs(file));
        }

        [Category("!OpenFile"), Description("Occurs when a file selection is changed")]
        public event SelectionChangedEventHandler SelectionChanged;

        private void RaiseSelectionChanged(string[] files)
        {
            if (null != SelectionChanged)
                SelectionChanged(this, new SelectionChangedEventArgs(files));
        }

        #endregion

        #region Properties

        [Category("Selection"), Description("Current selected file.")]
        public string SelectedFile
        {
            get
            {
                return Misc.SelectedFile;
            }
        }

        [Category("Selection"), Description("Current selected files.")]
        public string[] SelectedFiles
        {
            get
            {
                return Misc.SelectedFiles;
            }
        }

        [Category("Settings"), DisplayName("Localization"), Description("Provides localization settings."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public OpenFileLocalization Localization { get; private set; }

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

        private bool IsAlreadyLoaded { get; set; }

        private bool IsInDesignMode
        {
            get
            {
                System.ComponentModel.Design.IDesignerHost host;
                if (Site != null)
                {
                    host = Site.GetService(typeof(System.ComponentModel.Design.IDesignerHost)) as System.ComponentModel.Design.IDesignerHost;
                    if (host != null)
                    {
                        if (host.RootComponent.Site.DesignMode)
                            return true;
                        else
                            return false;
                    }
                }
                return false;
            }
        }

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
            if (!IsAlreadyLoaded && !IsInDesignMode)
            { 
                return;
            }
            Console.WriteLine("ShowAll");
            bool[] expanded = GetExpandedStates();
            TreeNode node = TryGetCurrentSelectedRootNode();
            ClearView();
            ShowDesktop(expanded[0]);
            ShowMyComputer(expanded[1]);
            ShowMyDocuments(expanded[2]);
            ShowSpecialFolders(expanded[3]);
            ShowTemplateFolders(expanded[4]);
            ApplyViewSettings();
            SelectFirstAvailable(node);
            DefaultableSettings settings = GetCurrentSettings();
            if (null != settings)
                ListView1.MultiSelect = settings.GetRuntimeValue("AllowMultipleSelect");
        }
         
        internal void ShowDesktop(bool expanded = false)
        {
            FileSystemInfo fsi = FileSystemHandler.GetDesktop();
            NodeDesktop.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeDesktop);
            if (settings.HasAllowedSubFolders(fsi))
            {
                if (expanded)
                {
                    NodeDesktop.Nodes.Add("?:");
                    NodeDesktop.Expand();
                }
                else
                {
                    NodeDesktop.Nodes.Add("?:");
                    NodeDesktop.Collapse();
                }
            }
        }

        internal void ShowMyComputer(bool expanded = false)
        {
            FileSystemInfo fsi = FileSystemHandler.GetMyComputer();
            NodeMyComputer.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeMyComputer);
            if (settings.HasAllowedSubFolders(fsi))
            {
                if (expanded)
                {
                    NodeMyComputer.Nodes.Add("?:");
                    NodeMyComputer.Expand();
                }
                else
                {
                    NodeMyComputer.Nodes.Add("?:");
                    NodeMyComputer.Collapse();
                }
            }
        }

        internal void ShowMyDocuments(bool expanded = false)
        {
            FileSystemInfo fsi = FileSystemHandler.GetMyDocuments();
            NodeMyDocuments.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeMyDocuments);
            if (settings.HasAllowedSubFolders(fsi))
            {
                if (expanded)
                {
                    NodeMyDocuments.Nodes.Add("?:");
                    NodeMyDocuments.Expand();
                }
                else
                {
                    NodeMyDocuments.Nodes.Add("?:");
                    NodeMyComputer.Collapse();
                }
            }
        }

        internal void ShowSpecialFolders(bool expanded = false)
        {
            FileSystemInfo fsi = FileSystemHandler.GetSpecialFolders();
            NodeSpecialFolders.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeSpecialFolders);
            if (settings.HasAllowedSubFolders(fsi))
            {
                if (expanded)
                {
                    NodeSpecialFolders.Nodes.Add("?:");
                    NodeSpecialFolders.Expand();
                }
                else
                {
                    NodeSpecialFolders.Nodes.Add("?:");
                    NodeSpecialFolders.Collapse();
                }
            }
        }

        internal void ShowTemplateFolders(bool expanded = false)
        {
            FileSystemInfo fsi = FileSystemHandler.GetTemplateFolders(TemplateFolders.FolderTemplates.ToArray());
            NodeTemplateFolders.Tag = fsi;
            DefaultableSettings settings = GetCurrentSettings(NodeTemplateFolders);
            if (settings.HasAllowedSubFolders(fsi))
            {
                if (expanded)
                {
                    NodeTemplateFolders.Nodes.Add("?:");
                    NodeTemplateFolders.Expand();
                }
                else
                {
                    NodeTemplateFolders.Nodes.Add("?:");
                    NodeTemplateFolders.Collapse();
                }
            }
        }

        #endregion

        #region ApplySettings Methods

        private void ApplySelectedCategory()
        {
            switch (Misc.SelectedCategory)
            {
                case RootCategory.Desktop:
                    if (TreeView1.Nodes.Contains(NodeDesktop))
                    TreeView1.SelectedNode = NodeDesktop;
                    break;
                case RootCategory.MyComputer:
                    if (TreeView1.Nodes.Contains(NodeMyComputer))
                        TreeView1.SelectedNode = NodeMyComputer;
                    break;
                case RootCategory.MyDocuments:
                    if (TreeView1.Nodes.Contains(NodeMyDocuments))
                        TreeView1.SelectedNode = NodeMyDocuments;
                    break;
                case RootCategory.SpecialFolders:
                    if (TreeView1.Nodes.Contains(NodeSpecialFolders))
                        TreeView1.SelectedNode = NodeSpecialFolders;
                    break;
                case RootCategory.TemplateFolders:
                    if (TreeView1.Nodes.Contains(NodeTemplateFolders))
                        TreeView1.SelectedNode = NodeTemplateFolders;
                    break;
                default:
                    break;
            }
        }

        private void ApplyLocalizationSettings(string propertyName = null)
        {
            if (null == propertyName)
                propertyName = "";

            switch (propertyName)
            {
                case "Desktop":
                    NodeDesktop.Text = Localization.Desktop;
                    break;
                case "MyMachine":
                    NodeMyComputer.Text = Localization.MyMachine;
                    break;
                case "MyDocuments":
                    NodeMyDocuments.Text = Localization.MyDocuments;
                    break;
                case "SpecialFolders":
                    NodeSpecialFolders.Text = Localization.SpecialFolders;
                    break;
                case "TemplateFolders":
                    NodeTemplateFolders.Text = Localization.TemplateFolders;
                    break;
                case "LabelFileName":
                    labelFileName.Text = Localization.LabelFileName;
                    break;
                case "LabelFileFilter":
                    labelFileType.Text = Localization.LabelFileFilter;
                    break;
                case "LabelLargeIconView":
                    StripButtonViewLargeIcon.Text = Localization.LabelLargeIconView;
                    break;
                case "LabelSmallIconView":
                    StripButtonViewSmallIcon.Text = Localization.LabelSmallIconView;
                    break;
                case "LabelDetailsView":
                    StripButtonViewDetails.Text = Localization.LabelDetailsView;
                    break;
                case "LabelCreateDirectory":
                    StripButtonCreateDirectory.Text = Localization.LabelCreateDirectory;
                    break;
                case "LabelDeleteDirectory":
                    StripButtonDeleteDirectory.Text = Localization.LabelDeleteDirectory;
                    break;
                case "LabelDeleteFile":
                    StripButtonDeleteFile.Text = Localization.LabelDeleteFile;
                    break;
                case "LabelGoUpward":
                    StripButtonGoUpward.Text = Localization.LabelGoUpward;
                    break;
                case "LabelGoUndo":
                    StripButtonGoUndo.Text = Localization.LabelGoUndo;
                    break;
                case "LabelGoRedo":
                    StripButtonGoRedo.Text = Localization.LabelGoRedo;
                    break;
                default:
                    NodeDesktop.Text = Localization.Desktop;
                    NodeMyComputer.Text = Localization.MyMachine;
                    NodeMyDocuments.Text = Localization.MyDocuments;
                    NodeSpecialFolders.Text = Localization.SpecialFolders;
                    NodeTemplateFolders.Text = Localization.TemplateFolders;
                    labelFileName.Text = Localization.LabelFileName;
                    labelFileType.Text = Localization.LabelFileFilter;
                    StripButtonViewLargeIcon.Text = Localization.LabelLargeIconView;
                    StripButtonViewSmallIcon.Text = Localization.LabelSmallIconView;
                    StripButtonViewDetails.Text = Localization.LabelDetailsView;
                    StripButtonCreateDirectory.Text = Localization.LabelCreateDirectory;
                    StripButtonDeleteDirectory.Text = Localization.LabelDeleteDirectory;
                    StripButtonDeleteFile.Text = Localization.LabelDeleteFile;
                    StripButtonGoUpward.Text = Localization.LabelGoUpward;
                    StripButtonGoUndo.Text = Localization.LabelGoUndo;
                    StripButtonGoRedo.Text = Localization.LabelGoRedo;
                    break;
            }
        }

        private void ApplyMiscSettings(string propertyName = null)
        {
            if (null == propertyName)
                propertyName = string.Empty;

            switch (propertyName)
            {
                case "SelectedCategory":
                    ApplySelectedCategory();
                    break;
                case "FileFilter":
                    comboBoxFileTypes.DataSource = Misc.Filters;
                    break;
                case "ShowCategoryPanel":
                    splitContainer1.Panel1Collapsed = !Misc.ShowCategoryPanel;
                    break;
                case "ShowFilePanel":
                    if (Misc.ShowFilePanel)
                    {
                        panelFileNameFilter.Visible = true;
                        splitContainer1.Height = this.Height - panelFileNameFilter.Height;
                    }
                    else
                    {
                        panelFileNameFilter.Visible = false;
                        splitContainer1.Height = this.Height;
                    }
                    break;
                case "CategoryPanelWidth":
                    splitContainer1.SplitterDistance = Misc.CategoryPanelWidth;
                    break;
                default:
                    ApplySelectedCategory();
                    if (Misc.ShowFilePanel)
                    {
                        panelFileNameFilter.Visible = true;
                        splitContainer1.Height = this.Height - panelFileNameFilter.Height;
                    }
                    else
                    {
                        panelFileNameFilter.Visible = false;
                        splitContainer1.Height = this.Height;
                    }
                    comboBoxFileTypes.DataSource = Misc.Filters;
                    splitContainer1.Panel1Collapsed = !Misc.ShowCategoryPanel;
                    splitContainer1.SplitterDistance = Misc.CategoryPanelWidth;
                    break;
            }
            
            TreeView1_AfterSelect(TreeView1, new TreeViewEventArgs(TreeView1.SelectedNode));
        }
        
        private void ApplyDefaultableSettings(string propertyName = null)
        {
            UdpateRootCategoryVisibility(NodeDesktop, Desktop, GetInsertNodeIndex(RootCategory.Desktop));
            UdpateRootCategoryVisibility(NodeMyComputer, MyComputer, GetInsertNodeIndex(RootCategory.MyComputer));
            UdpateRootCategoryVisibility(NodeMyDocuments, MyDocuments, GetInsertNodeIndex(RootCategory.MyDocuments));
            UdpateRootCategoryVisibility(NodeSpecialFolders, SpecialFolders, GetInsertNodeIndex(RootCategory.SpecialFolders));
            UdpateRootCategoryVisibility(NodeTemplateFolders, TemplateFolders, GetInsertNodeIndex(RootCategory.TemplateFolders));            
            ShowAll();
            UdpateRootCategoryExpandedState(NodeDesktop, Desktop);
            UdpateRootCategoryExpandedState(NodeMyComputer, MyComputer);
            UdpateRootCategoryExpandedState(NodeMyDocuments, MyDocuments);
            UdpateRootCategoryExpandedState(NodeSpecialFolders, SpecialFolders);
            UdpateRootCategoryExpandedState(NodeTemplateFolders, TemplateFolders);
        }

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

        #region Update UI Methods

        private void UdpateRootCategoryExpandedState(TreeNode node, DefaultableSettings settings)
        {
            if (settings.Expanded && node.IsExpanded == false)
                node.Expand();
            if (!settings.Expanded && node.IsExpanded == true)
                node.Collapse();
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

        private void ShowCurrentNodeInListView(TreeNode node, FileSystemInfo fsInfo)
        {
            DefaultableSettings settings = GetCurrentSettings(node);
            ListView1.MultiSelect = settings.GetRuntimeValue("AllowMultipleSelect");
            if (settings.GetRuntimeValue("AllowBrowseFolders"))
            {
                foreach (DrvInfo item in fsInfo.Drives)
                {
                    if (settings.AllowShowDrive(item))
                    {
                        ListViewItem viewItem = ListView1.Items.Add(item.Name + item.Label);
                        viewItem.Tag = item;
                        viewItem.ImageIndex = CalculateDriveImageIndexListView(item);
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
                    viewItem.ImageIndex = CalculateFileImageIndex(item.Name);
                }
            }

            StripButtonCreateDirectory.Enabled = settings.CanCreateFolders(fsInfo);
            StripButtonDeleteDirectory.Enabled = false;
            StripButtonDeleteFile.Enabled = false;
        }

        #endregion

        #region Other Methods

        private bool[] GetExpandedStates()
        {
            List<bool> list = new List<bool>();
            list.Add(Desktop.Expanded);
            list.Add(MyComputer.Expanded);
            list.Add(MyDocuments.Expanded);
            list.Add(SpecialFolders.Expanded);
            list.Add(TemplateFolders.Expanded);
            return list.ToArray();
        }

        private void SelectFirstAvailable(TreeNode node = null)
        {
            if (node != null && TreeView1.Nodes.Contains(node))
            {
                TreeView1.SelectedNode = node;
                TreeView1_AfterSelect(TreeView1, new TreeViewEventArgs(TreeView1.SelectedNode));
                return;
            }

            bool treeViewNeedHelp = false;
            if (TreeView1.Nodes.Contains(NodeDesktop))
            {
                if (TreeView1.SelectedNode == NodeDesktop)
                    treeViewNeedHelp = true;
                TreeView1.SelectedNode = NodeDesktop;
            }
            else if (TreeView1.Nodes.Contains(NodeMyComputer))
            {
                if (TreeView1.SelectedNode == NodeDesktop)
                    treeViewNeedHelp = true;
                TreeView1.SelectedNode = NodeMyComputer;
            }
            else if (TreeView1.Nodes.Contains(NodeMyDocuments))
            {
                if (TreeView1.SelectedNode == NodeDesktop)
                    treeViewNeedHelp = true;
                TreeView1.SelectedNode = NodeMyDocuments;
            }
            else if (TreeView1.Nodes.Contains(NodeSpecialFolders))
            {
                if (TreeView1.SelectedNode == NodeDesktop)
                    treeViewNeedHelp = true;
                TreeView1.SelectedNode = NodeSpecialFolders;
            }
            else if (TreeView1.Nodes.Contains(NodeTemplateFolders))
            {
                if (TreeView1.SelectedNode == NodeDesktop)
                    treeViewNeedHelp = true;
                TreeView1.SelectedNode = NodeTemplateFolders;
            }
            if (treeViewNeedHelp)
                TreeView1_AfterSelect(TreeView1, new TreeViewEventArgs(TreeView1.SelectedNode));
        }

        private string[] GetSelectedFiles()
        {
            List<string> list = new List<string>();
            foreach (ListViewItem item in ListView1.SelectedItems)
            {
                FileSystemInfo fsi = item.Tag as FileSystemInfo;
                if (fsi is FiInfo)
                    list.Add(fsi.Path);
            }
            return list.ToArray();
        }

        private string CreateSelectedFilesString(List<string> list)
        {
            if (list.Count == 0)
                return "";
            if (list.Count == 1)
                return System.IO.Path.GetFileName(list[0]);
            string result = "";
            foreach (var item in list)
            {
                string fileName = System.IO.Path.GetFileName(item);
                result += fileName + ";";
            }
            result = result.Substring(0, result.Length - 1);
            return result;
        }

        private TreeNode TryGetCurrentSelectedRootNode()
        {
            if (null == TreeView1.SelectedNode)
                return null;
            TreeNode selectedNode = TreeView1.SelectedNode;
            if (selectedNode.Parent == null)
            {
                switch (selectedNode.Name)
                {
                    case "NodeDesktop":
                    case "NodeMyComputer":
                    case "NodeMyDocuments":
                    case "NodeSpecialFolders":
                    case "NodeTemplateFolders":
                        return selectedNode;
                }
            }
            return null;
        }

        private TreeNode GetCurrentSelectedRootNode()
        {
            if (null == TreeView1.SelectedNode)
                return null;
            TreeNode selectedNode = TreeView1.SelectedNode;

            while (selectedNode != null)
            {
                if (selectedNode.Parent == null)
                {
                    switch (selectedNode.Name)
                    {
                        case "NodeDesktop":
                        case "NodeMyComputer":
                        case "NodeMyDocuments":
                        case "NodeSpecialFolders":
                        case "NodeTemplateFolders":
                            return selectedNode;
                    }
                }
                selectedNode = selectedNode.Parent;
            }

            return null;
        }

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
       
        private int GetInsertNodeIndex(RootCategory category)
        {
            int targetIndex = 0;
            switch (category)
            {
                case RootCategory.MyComputer:
                    if (TreeView1.Nodes.Contains(NodeDesktop))
                        targetIndex++;
                    break;
                case RootCategory.MyDocuments:
                    if (TreeView1.Nodes.Contains(NodeDesktop))
                        targetIndex++;
                    if (TreeView1.Nodes.Contains(NodeMyComputer))
                        targetIndex++;
                    break;
                case RootCategory.SpecialFolders:
                    if (TreeView1.Nodes.Contains(NodeDesktop))
                        targetIndex++;
                    if (TreeView1.Nodes.Contains(NodeMyComputer))
                        targetIndex++;
                    if (TreeView1.Nodes.Contains(NodeMyDocuments))
                        targetIndex++;
                    break;
                case RootCategory.TemplateFolders:
                    if (TreeView1.Nodes.Contains(NodeDesktop))
                        targetIndex++;
                    if (TreeView1.Nodes.Contains(NodeMyComputer))
                        targetIndex++;
                    if (TreeView1.Nodes.Contains(NodeMyDocuments))
                        targetIndex++;
                    if (TreeView1.Nodes.Contains(NodeSpecialFolders))
                        targetIndex++;
                    break;
                default:
                    return 0;
            }
            return targetIndex;
        }

        private int CalculateDriveImageIndexTreeView(DrvInfo item)
        {
            switch (item.Type)
            {
                case System.IO.DriveType.CDRom:
                    return item.IsReady ? 8 : 9;
                case System.IO.DriveType.Network:
                    return item.IsReady ? 10 : 11;
                default:
                    return item.IsReady ? 6 : 7;
            }
        }

        private int CalculateDriveImageIndexListView(DrvInfo item)
        {
            switch (item.Type)
            {
                case System.IO.DriveType.CDRom:
                    return item.IsReady ? 4 : 5;
                case System.IO.DriveType.Network:
                    return item.IsReady ? 6 : 7;
                default:
                    return item.IsReady ? 2 : 3;
            }
        }

        private void GotoSubNode(TreeNode node, string subNodeName)
        {
            if (!node.IsExpanded)
                node.Expand();

            foreach (TreeNode item in node.Nodes)
            {
                FileSystemInfo fsi = item.Tag as FileSystemInfo;
                if (fsi.Name == subNodeName)
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
                    TreeNode subNode = node.Nodes.Add(item.Name + item.Label);
                    subNode.Tag = item;

                    int drvImageIndex = CalculateDriveImageIndexTreeView(item);
                    subNode.ImageIndex = drvImageIndex;
                    subNode.SelectedImageIndex = drvImageIndex;

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
                    subNode.SelectedImageIndex = 12;

                    if (GetCurrentSettings(subNode).HasAllowedSubFolders(item))
                        subNode.Nodes.Add("?:");
                }
            }
        }
         
        private int CalculateFileImageIndex(string file)
        {
            if (String.IsNullOrWhiteSpace(file))
                return 0;
            string fileExtension = System.IO.Path.GetExtension(file).ToLower().Trim();
              if (String.IsNullOrWhiteSpace(fileExtension))
                return 0;
            switch (fileExtension)
            {
                case ".avi":
                case ".mpg":
                case ".wmf":
                case ".wsf":
                case ".divx":
                case ".xdiv":
                case ".flv":
                    return 8;
                case ".bmp":
                    return 9;
                case ".c":
                case ".h":
                    return 10;
                case ".cpp":
                    return 11;
                case ".cs":
                    return 12;
                case ".sql":
                    return 13;
                case ".exe":
                    return 14;
                case ".mp3":
                case ".mp4":
                case ".ogg":
                case ".wav":
                    return 15;
                case ".php":
                    return 16;
                case ".jpg":
                case ".jpe":
                case ".jpeg":
                case ".png":
                case ".dib":
                case ".gif":
                case ".tif":
                case ".tiff":
                    return 17;
                case ".rtf":
                    return 18;
                case ".txt":
                    return 19;
                case ".vb":
                    return 20;
                case ".vcf":
                    return 21;
                case ".bat":
                    return 22;
                default:
                    return 0;
            }
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
        
        private ListViewItem GetDirectoryListViewItem(string path)
        {
            foreach (ListViewItem item in ListView1.Items)
            {
                FileSystemInfo fsi = item.Tag as FileSystemInfo;
                if (fsi is FolderInfo)
                {
                    if (path.Equals(fsi.Path, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                }
            }
            return null;
        }

        private ListViewItem GetFileListViewItem(string path)
        {
            foreach (ListViewItem item in ListView1.Items)
            {
                FileSystemInfo fsi = item.Tag as FileSystemInfo;
                if (fsi is FiInfo)
                {
                    if (path.Equals(fsi.Path, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                }
            }
            return null;
        }
         
        #endregion

        #region Directory/File Methods

        private string GetNewFolderFullPathNameForCurrentSelectedDirectory()
        {
            FolderInfo fiInfo = TreeView1.SelectedNode.Tag as FolderInfo;
            if (null == fiInfo)
                return null;
            int position = fiInfo.Path.LastIndexOf("\\");
            if (position < 0)
                return null;

            string dirName = fiInfo.Path;

            for (int i = 0; i < 1024; i++)
            {
                string newDirectory = Localization.NewDirectoryName;
                string newDirectoryName = newDirectory;
                string fullPath = System.IO.Path.Combine(dirName, newDirectoryName);
                if (i == 0)
                {
                    if (System.IO.Directory.Exists(fullPath))
                        continue;
                    else
                        return fullPath;
                }
                else
                {
                    fullPath = String.Format("{0}\\{1}({2})", dirName, newDirectoryName, i);
                    if (System.IO.Directory.Exists(fullPath))
                        continue;
                    else
                        return fullPath;
                }
            }

            return String.Empty;
        }

        private void AddNewDirectory(string folder)
        {
            var settings = GetCurrentSettings();
            if (settings.GetRuntimeValue("AllowBrowseFolders") == false)
                return;
            string folderName = System.IO.Path.GetFileNameWithoutExtension(folder);
            FolderInfo info = new FolderInfo(FileSystemHandler, folderName, folder);
            ListViewItem item = ListView1.Items.Add(info.Name);
            if (TreeView1.SelectedNode != null)
            {
                if (TreeView1.SelectedNode != null && TreeView1.SelectedNode.Nodes[0].Text != "?:")
                {
                    TreeNode node = TreeView1.SelectedNode.Nodes.Insert(0, folderName);
                    node.Tag = new FolderInfo(FileSystemHandler, folderName, folder);
                }

            }
            item.ImageIndex = 1;
            item.Tag = info;
        }

        private void CreateDirectory()
        {
            string directoryName = String.Empty;
            try
            {
                directoryName = GetNewFolderFullPathNameForCurrentSelectedDirectory();
                if (!String.IsNullOrWhiteSpace(directoryName))
                {
                    System.IO.Directory.CreateDirectory(directoryName);
                    AddNewDirectory(directoryName);
                }
            }
            catch (Exception exception)
            {
                OnDirectoryCreateError(directoryName, exception);
            }
        }

        private void DeleteSelectedFiles()
        {
            string[] files = GetCurrentFileSelection();
            if (files.Length == 0)
                return;
            if (Misc.AskBeforeDelete)
            {
                bool cancel = false;
                ConfirmDeleteFiles(files, ref cancel);
                if (cancel)
                    return;
            }

            string[] deletedFiles = DeleteFiles(files);
            foreach (var item in deletedFiles)
            {
                ListViewItem lviItem = GetFileListViewItem(item);
                if (null != lviItem)
                    ListView1.Items.Remove(lviItem);
            }
        }

        private void DeleteSelectedFolders()
        {
            string[] folders = GetCurrentDirectorySelection();
            if (folders.Length == 0)
                return;
            if (Misc.AskBeforeDelete)
            {
                bool cancel = false;
                ConfirmDeleteDirectories(folders, ref cancel);
                if (cancel)
                    return;
            }

            string[] deletedFolders = DeleteDirectories(folders);
            foreach (var item in deletedFolders)
            {
                ListViewItem lviItem = GetDirectoryListViewItem(item);
                if (null != lviItem)
                    ListView1.Items.Remove(lviItem);
                TreeNode currentNode = TreeView1.SelectedNode;
                if (null != currentNode)
                    DeleteDirectorySubNode(currentNode, item);
            }
        }

        private void DeleteDirectorySubNode(TreeNode node, string path)
        {
            if (null == node)
                return;
            foreach (TreeNode item in node.Nodes)
            {
                FileSystemInfo fsi = item.Tag as FileSystemInfo;
                if (fsi is FolderInfo)
                {
                    if (path.Equals(fsi.Path, StringComparison.InvariantCultureIgnoreCase))
                    {
                        node.Nodes.Remove(item);
                        return;
                    }
                }
            }
        }

        private string[] DeleteFiles(string[] files)
        {
            List<string> list = new List<string>();
            string currentFile = string.Empty;
            foreach (var item in files)
            {
                try
                {
                    System.IO.File.Delete(item);
                    list.Add(item);
                }
                catch (Exception exception)
                {
                    bool handled = OnFileDeleteError(currentFile, exception);
                    if (handled)
                        continue;
                    else
                        return list.ToArray();
                }
            }
            return list.ToArray();
        }

        private string[] DeleteDirectories(string[] directories)
        {
            List<string> list = new List<string>();
            string currentFile = string.Empty;
            foreach (var item in directories)
            {
                try
                {
                    System.IO.Directory.Delete(item, true);
                    list.Add(item);
                }
                catch (Exception exception)
                {
                    bool handled = OnDirectoryDeleteError(currentFile, exception);
                    if (handled)
                        continue;
                    else
                        return list.ToArray();
                }
            }
            return list.ToArray();
        }

        private string[] GetCurrentFileSelection()
        {
            List<string> list = new List<string>();
            foreach (ListViewItem item in ListView1.SelectedItems)
            {
                FileSystemInfo fsi = item.Tag as FileSystemInfo;
                if (fsi is FiInfo)
                    list.Add(fsi.Path);
            }
            return list.ToArray();
        }

        private string[] GetCurrentDirectorySelection()
        {
            List<string> list = new List<string>();
            foreach (ListViewItem item in ListView1.SelectedItems)
            {
                FileSystemInfo fsi = item.Tag as FileSystemInfo;
                if (fsi is FolderInfo)
                    list.Add(fsi.Path);
            }
            return list.ToArray();
        }

        #endregion

        #region Virtual Methods

        protected virtual bool OnFileDeleteError(string file, Exception exception)
        {
            return false;
        }

        protected virtual void OnDirectoryCreateError(string directory, Exception exception)
        {

        }

        protected virtual bool OnDirectoryDeleteError(string file, Exception exception)
        {
            return false;
        }

        protected virtual void ConfirmDeleteDirectories(string[] directories, ref bool cancel)
        {
            DialogResult dr = MessageBox.Show(this, Localization.AskBeforeDeleteDirectoryMessage, Localization.AskBeforeDeleteDirectoryHeader, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr != DialogResult.Yes)
                cancel = true;
        }

        protected virtual void ConfirmDeleteFiles(string[] files, ref bool cancel)
        {
            DialogResult dr = MessageBox.Show(this, Localization.AskBeforeDeleteFileMessage, Localization.AskBeforeDeleteFileHeader, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr != DialogResult.Yes)
                cancel = true;
        }

        #endregion

        #region Settings Trigger

        private void LocalizationSettings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            ApplyLocalizationSettings(e.PropertyName);
        }

        private void MiscSettings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
             ApplyMiscSettings(e.PropertyName);
        }

        private void DefaultableSettings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            ApplyDefaultableSettings(e.PropertyName);
        }

        #endregion

        #region UserControl Trigger

        private void OpenFilePanel_Load(object sender, EventArgs e)
        {
            IsAlreadyLoaded = true;
            ShowAll();
        }

        private void OpenFilePanel_Resize(object sender, EventArgs e)
        {
            colHeaderSize.Width = 100;
            colHeaderName.Width = ListView1.Width - (colHeaderSize.Width + 10);
        }

        #endregion
        
        #region TreeView Trigger

        private void TreeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            FileSystemInfo fsInfo = e.Node.Tag as FileSystemInfo;
            if (null == fsInfo)
                return;
            if (fsInfo is FolderInfo & e.Node.Parent != null)
            {
                e.Node.ImageIndex = 13;
                e.Node.SelectedImageIndex = 13;
            }

            DefaultableSettings settings = GetCurrentSettings(e.Node);
            if (!settings.HasAllowedSubFolders(fsInfo))
            {
                e.Node.Nodes.Clear();
                return;
            }

            ExpandNode(e.Node, fsInfo);
        }

        private void TreeView1_AfterExpand(object sender, TreeViewEventArgs e)
        {
            if (e.Node == NodeDesktop)
                Desktop.SetExpanded(true);
            else if (e.Node == NodeMyComputer)
                MyComputer.SetExpanded(true);
            else if (e.Node == NodeMyDocuments)
                MyDocuments.SetExpanded(true);
            else if (e.Node == NodeSpecialFolders)
                SpecialFolders.SetExpanded(true);
            else if (e.Node == NodeTemplateFolders)
                TemplateFolders.SetExpanded(true);
        }

        private void TreeView1_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (!IsInDesignMode && !IsAlreadyLoaded)
            {
                switch (GetCurrentSelectedRootCategory(e.Node))
                {
                    case RootCategory.Desktop:
                        if (Desktop.Expanded)
                            e.Cancel = true;
                        break;
                    case RootCategory.MyComputer:
                        if (MyComputer.Expanded)
                            e.Cancel = true;
                        break;
                    case RootCategory.MyDocuments:
                        if (MyDocuments.Expanded)
                            e.Cancel = true;
                        break;
                    case RootCategory.SpecialFolders:
                        if (SpecialFolders.Expanded)
                            e.Cancel = true;
                        break;
                    case RootCategory.TemplateFolders:
                        if (TemplateFolders.Expanded)
                            e.Cancel = true;
                        break;
                }
            }
        }

        private void TreeView1_AfterCollapse(object sender, TreeViewEventArgs e)
        {
            if (null == e.Node)
                return;
            FileSystemInfo fsInfo = e.Node.Tag as FileSystemInfo;
            if (null == fsInfo)
                return;
            if (fsInfo is FolderInfo && e.Node.Parent != null)
            {
                e.Node.ImageIndex = 12;
                e.Node.SelectedImageIndex = 12;
            }

            if (e.Node == NodeDesktop)
                Desktop.SetExpanded(false);
            else if (e.Node == NodeMyComputer)
                MyComputer.SetExpanded(false);
            else if (e.Node == NodeMyDocuments)
                MyDocuments.SetExpanded(false);
            else if (e.Node == NodeSpecialFolders)
                SpecialFolders.SetExpanded(false);
            else if (e.Node == NodeTemplateFolders)
                TemplateFolders.SetExpanded(false);
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
            Misc.SetSelectedCategory(GetCurrentSelectedRootCategory());
            Misc.SetSelectedFiles(new string[0]);
            textBoxSelectedFiles.Text = String.Empty;
            RaiseSelectionChanged(new string[0]);
        }

        #endregion

        #region ListView Trigger

        private void ListView1_DoubleClick(object sender, EventArgs e)
        {
            if (ListView1.SelectedItems.Count == 0)
                return;

            FileSystemInfo fsInfo = ListView1.SelectedItems[0].Tag as FileSystemInfo;
            if (fsInfo is FiInfo)
            {
                if (Misc.FireSelectionChangedInsteadOfDoubleClick)
                    RaiseSelectionChanged(new string[] { fsInfo.Path });
                else 
                    RaiseFileDoubleClick(fsInfo.Path);
            }
            else
            {
                GotoSubNode(TreeView1.SelectedNode, fsInfo.Name);
            }
        }

        private void ListView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            UpdateToolStripButtons();
            string[] files = GetSelectedFiles();
            RaiseSelectionChanged(files);
        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            foreach (ListViewItem item in ListView1.SelectedItems)
            {
                FileSystemInfo fsi = item.Tag as FileSystemInfo;
                if (fsi is FiInfo)
                    list.Add(fsi.Path);
            }
            Misc.SetSelectedFiles(list.ToArray());
            textBoxSelectedFiles.Text = CreateSelectedFilesString(list); ;
            UpdateToolStripButtons();
            string[] files = GetSelectedFiles();
            RaiseSelectionChanged(files);
        }

        #endregion

        #region Strip Trigger

        private void ToolStripContainer1_ContentPanel_SizeChanged(object sender, EventArgs e)
        {
            splitContainer1.Size = new Size(splitContainer1.Parent.Width - 2, splitContainer1.Parent.Height - 68);
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
            OpenFilePanel_Resize(this, new EventArgs());
        }

        private void StripButtonCreateDirectory_Click(object sender, EventArgs e)
        {
            CreateDirectory();
        }

        private void StripButtonDeleteDirectory_Click(object sender, EventArgs e)
        {
            DeleteSelectedFolders();
        }

        private void StripButtonDeleteFile_Click(object sender, EventArgs e)
        {
            DeleteSelectedFiles();
        }

        #endregion

        #region FilePanel Trigger

        private void comboBoxFileTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!IsInDesignMode && !IsAlreadyLoaded)
                return;
            Misc.CurrentFilter = comboBoxFileTypes.SelectedItem as FileFilterItem;
            if (null != TreeView1.SelectedNode)
            {
                ListView1.Items.Clear();
                ShowCurrentNodeInListView(TreeView1.SelectedNode, TreeView1.SelectedNode.Tag as FileSystemInfo);
            }
        }

        #endregion

        #region SplitContainer Trigger

        private void splitContainer1_Panel1_Resize(object sender, EventArgs e)
        {
            // docking is bugy
            if (null != Misc)
                Misc.SetCategoryPanelWidth(splitContainer1.Panel1.Width);
        }

        #endregion
    }
}
