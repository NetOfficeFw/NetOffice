namespace NOTools.FileSystemDialogs
{
    partial class OpenFilePanel
    {
        /// <summary> 
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary> 
        /// Erforderliche Methode für die Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Desktop", 1, 1);
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("My Computer", 2, 2);
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("My Documents", 0, 0);
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Special Folders", 4, 4);
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Template Folders", 5, 5);
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OpenFilePanel));
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem("Item1", 1);
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem("Item2", 1);
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem("Item3", 1);
            System.Windows.Forms.ListViewItem listViewItem4 = new System.Windows.Forms.ListViewItem("Item4", 1);
            this.labelFileType = new System.Windows.Forms.Label();
            this.labelFileName = new System.Windows.Forms.Label();
            this.comboBoxFileTypes = new System.Windows.Forms.ComboBox();
            this.textBoxSelectedFiles = new System.Windows.Forms.TextBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.TreeView1 = new System.Windows.Forms.TreeView();
            this.imageListLeft = new System.Windows.Forms.ImageList(this.components);
            this.ListView1 = new System.Windows.Forms.ListView();
            this.colHeaderName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colHeaderSize = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageListLarge = new System.Windows.Forms.ImageList(this.components);
            this.imageListSmall = new System.Windows.Forms.ImageList(this.components);
            this.toolStripContainer1 = new System.Windows.Forms.ToolStripContainer();
            this.panelFileNameFilter = new System.Windows.Forms.Panel();
            this.toolStrip2 = new System.Windows.Forms.ToolStrip();
            this.StripButtonViewLargeIcon = new System.Windows.Forms.ToolStripButton();
            this.StripButtonViewSmallIcon = new System.Windows.Forms.ToolStripButton();
            this.StripButtonViewDetails = new System.Windows.Forms.ToolStripButton();
            this.ToolStrip1 = new System.Windows.Forms.ToolStrip();
            this.StripButtonGoUpward = new System.Windows.Forms.ToolStripButton();
            this.StripButtonGoUndo = new System.Windows.Forms.ToolStripButton();
            this.StripButtonGoRedo = new System.Windows.Forms.ToolStripButton();
            this.StripButtonCreateDirectory = new System.Windows.Forms.ToolStripButton();
            this.StripButtonDeleteDirectory = new System.Windows.Forms.ToolStripButton();
            this.StripButtonDeleteFile = new System.Windows.Forms.ToolStripButton();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.toolStripContainer1.ContentPanel.SuspendLayout();
            this.toolStripContainer1.RightToolStripPanel.SuspendLayout();
            this.toolStripContainer1.SuspendLayout();
            this.panelFileNameFilter.SuspendLayout();
            this.toolStrip2.SuspendLayout();
            this.ToolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelFileType
            // 
            this.labelFileType.AutoSize = true;
            this.labelFileType.Location = new System.Drawing.Point(9, 35);
            this.labelFileType.Name = "labelFileType";
            this.labelFileType.Size = new System.Drawing.Size(49, 13);
            this.labelFileType.TabIndex = 9;
            this.labelFileType.Text = "Dateityp:";
            // 
            // labelFileName
            // 
            this.labelFileName.AutoSize = true;
            this.labelFileName.Location = new System.Drawing.Point(9, 9);
            this.labelFileName.Name = "labelFileName";
            this.labelFileName.Size = new System.Drawing.Size(61, 13);
            this.labelFileName.TabIndex = 8;
            this.labelFileName.Text = "Dateiname:";
            // 
            // comboBoxFileTypes
            // 
            this.comboBoxFileTypes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxFileTypes.DisplayMember = "Name";
            this.comboBoxFileTypes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxFileTypes.FormattingEnabled = true;
            this.comboBoxFileTypes.Location = new System.Drawing.Point(75, 32);
            this.comboBoxFileTypes.Name = "comboBoxFileTypes";
            this.comboBoxFileTypes.Size = new System.Drawing.Size(506, 21);
            this.comboBoxFileTypes.TabIndex = 7;
            this.comboBoxFileTypes.SelectedIndexChanged += new System.EventHandler(this.comboBoxFileTypes_SelectedIndexChanged);
            // 
            // textBoxSelectedFiles
            // 
            this.textBoxSelectedFiles.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxSelectedFiles.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxSelectedFiles.Location = new System.Drawing.Point(75, 3);
            this.textBoxSelectedFiles.Name = "textBoxSelectedFiles";
            this.textBoxSelectedFiles.ReadOnly = true;
            this.textBoxSelectedFiles.Size = new System.Drawing.Size(503, 20);
            this.textBoxSelectedFiles.TabIndex = 6;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Margin = new System.Windows.Forms.Padding(1);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.TreeView1);
            this.splitContainer1.Panel1.Resize += new System.EventHandler(this.splitContainer1_Panel1_Resize);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.ListView1);
            this.splitContainer1.Size = new System.Drawing.Size(582, 323);
            this.splitContainer1.SplitterDistance = 230;
            this.splitContainer1.TabIndex = 10;
            // 
            // TreeView1
            // 
            this.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TreeView1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TreeView1.HideSelection = false;
            this.TreeView1.ImageIndex = 13;
            this.TreeView1.ImageList = this.imageListLeft;
            this.TreeView1.ItemHeight = 20;
            this.TreeView1.Location = new System.Drawing.Point(0, 0);
            this.TreeView1.Name = "TreeView1";
            treeNode1.ImageIndex = 1;
            treeNode1.Name = "NodeDesktop";
            treeNode1.SelectedImageIndex = 1;
            treeNode1.Text = "Desktop";
            treeNode2.ImageIndex = 2;
            treeNode2.Name = "NodeMyComputer";
            treeNode2.SelectedImageIndex = 2;
            treeNode2.Text = "My Computer";
            treeNode3.ImageIndex = 0;
            treeNode3.Name = "NodeMyDocuments";
            treeNode3.SelectedImageIndex = 0;
            treeNode3.Text = "My Documents";
            treeNode4.ImageIndex = 4;
            treeNode4.Name = "NodeSpecialFolders";
            treeNode4.SelectedImageIndex = 4;
            treeNode4.Text = "Special Folders";
            treeNode5.ImageIndex = 5;
            treeNode5.Name = "NodeTemplateFolders";
            treeNode5.SelectedImageIndex = 5;
            treeNode5.Text = "Template Folders";
            this.TreeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4,
            treeNode5});
            this.TreeView1.SelectedImageIndex = 0;
            this.TreeView1.Size = new System.Drawing.Size(230, 323);
            this.TreeView1.TabIndex = 1;
            this.TreeView1.BeforeCollapse += new System.Windows.Forms.TreeViewCancelEventHandler(this.TreeView1_BeforeCollapse);
            this.TreeView1.AfterCollapse += new System.Windows.Forms.TreeViewEventHandler(this.TreeView1_AfterCollapse);
            this.TreeView1.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.TreeView1_BeforeExpand);
            this.TreeView1.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.TreeView1_AfterExpand);
            this.TreeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.TreeView1_AfterSelect);
            // 
            // imageListLeft
            // 
            this.imageListLeft.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListLeft.ImageStream")));
            this.imageListLeft.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListLeft.Images.SetKeyName(0, "");
            this.imageListLeft.Images.SetKeyName(1, "");
            this.imageListLeft.Images.SetKeyName(2, "");
            this.imageListLeft.Images.SetKeyName(3, "drives.png");
            this.imageListLeft.Images.SetKeyName(4, "special_folders.png");
            this.imageListLeft.Images.SetKeyName(5, "template_folders.png");
            this.imageListLeft.Images.SetKeyName(6, "hard_drive.png");
            this.imageListLeft.Images.SetKeyName(7, "hard_drive_error.png");
            this.imageListLeft.Images.SetKeyName(8, "cd_drive.png");
            this.imageListLeft.Images.SetKeyName(9, "cd_drive_error.png");
            this.imageListLeft.Images.SetKeyName(10, "hard_drive_network.png");
            this.imageListLeft.Images.SetKeyName(11, "hard_drive_network_error.png");
            this.imageListLeft.Images.SetKeyName(12, "folder_closed.png");
            this.imageListLeft.Images.SetKeyName(13, "folder_open.png");
            // 
            // ListView1
            // 
            this.ListView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colHeaderName,
            this.colHeaderSize});
            this.ListView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ListView1.FullRowSelect = true;
            this.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.ListView1.HideSelection = false;
            this.ListView1.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3,
            listViewItem4});
            this.ListView1.LargeImageList = this.imageListLarge;
            this.ListView1.Location = new System.Drawing.Point(0, 0);
            this.ListView1.MultiSelect = false;
            this.ListView1.Name = "ListView1";
            this.ListView1.ShowGroups = false;
            this.ListView1.Size = new System.Drawing.Size(348, 323);
            this.ListView1.SmallImageList = this.imageListSmall;
            this.ListView1.TabIndex = 0;
            this.ListView1.UseCompatibleStateImageBehavior = false;
            this.ListView1.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.ListView1_ItemSelectionChanged);
            this.ListView1.SelectedIndexChanged += new System.EventHandler(this.ListView1_SelectedIndexChanged);
            this.ListView1.DoubleClick += new System.EventHandler(this.ListView1_DoubleClick);
            // 
            // colHeaderName
            // 
            this.colHeaderName.Text = "Name";
            this.colHeaderName.Width = 100;
            // 
            // colHeaderSize
            // 
            this.colHeaderSize.Text = "Size";
            this.colHeaderSize.Width = 100;
            // 
            // imageListLarge
            // 
            this.imageListLarge.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListLarge.ImageStream")));
            this.imageListLarge.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListLarge.Images.SetKeyName(0, "UnkownFile.png");
            this.imageListLarge.Images.SetKeyName(1, "Folder.png");
            this.imageListLarge.Images.SetKeyName(2, "hard_drive.png");
            this.imageListLarge.Images.SetKeyName(3, "hard_drive_error.png");
            this.imageListLarge.Images.SetKeyName(4, "cd_drive.png");
            this.imageListLarge.Images.SetKeyName(5, "cd_drive_error.png");
            this.imageListLarge.Images.SetKeyName(6, "hard_drive_network.png");
            this.imageListLarge.Images.SetKeyName(7, "hard_drive_network_error.png");
            this.imageListLarge.Images.SetKeyName(8, "avi.png");
            this.imageListLarge.Images.SetKeyName(9, "bmp.png");
            this.imageListLarge.Images.SetKeyName(10, "c.png");
            this.imageListLarge.Images.SetKeyName(11, "cplusplus.png");
            this.imageListLarge.Images.SetKeyName(12, "csharp.png");
            this.imageListLarge.Images.SetKeyName(13, "data.png");
            this.imageListLarge.Images.SetKeyName(14, "exe.png");
            this.imageListLarge.Images.SetKeyName(15, "mp3.png");
            this.imageListLarge.Images.SetKeyName(16, "php.png");
            this.imageListLarge.Images.SetKeyName(17, "png.png");
            this.imageListLarge.Images.SetKeyName(18, "rtf.png");
            this.imageListLarge.Images.SetKeyName(19, "txt.png");
            this.imageListLarge.Images.SetKeyName(20, "vbasic.png");
            this.imageListLarge.Images.SetKeyName(21, "vcf.png");
            this.imageListLarge.Images.SetKeyName(22, "bat.png");
            // 
            // imageListSmall
            // 
            this.imageListSmall.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListSmall.ImageStream")));
            this.imageListSmall.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListSmall.Images.SetKeyName(0, "UnkownFile.png");
            this.imageListSmall.Images.SetKeyName(1, "Folder.png");
            this.imageListSmall.Images.SetKeyName(2, "hard_drive.png");
            this.imageListSmall.Images.SetKeyName(3, "hard_drive_error.png");
            this.imageListSmall.Images.SetKeyName(4, "cd_drive.png");
            this.imageListSmall.Images.SetKeyName(5, "cd_drive_error.png");
            this.imageListSmall.Images.SetKeyName(6, "hard_drive_network.png");
            this.imageListSmall.Images.SetKeyName(7, "hard_drive_network_error.png");
            this.imageListSmall.Images.SetKeyName(8, "avi.png");
            this.imageListSmall.Images.SetKeyName(9, "bmp.png");
            this.imageListSmall.Images.SetKeyName(10, "c.png");
            this.imageListSmall.Images.SetKeyName(11, "cplusplus.png");
            this.imageListSmall.Images.SetKeyName(12, "csharp.png");
            this.imageListSmall.Images.SetKeyName(13, "data.png");
            this.imageListSmall.Images.SetKeyName(14, "exe.png");
            this.imageListSmall.Images.SetKeyName(15, "mp3.png");
            this.imageListSmall.Images.SetKeyName(16, "php.png");
            this.imageListSmall.Images.SetKeyName(17, "png.png");
            this.imageListSmall.Images.SetKeyName(18, "rtf.png");
            this.imageListSmall.Images.SetKeyName(19, "txt.png");
            this.imageListSmall.Images.SetKeyName(20, "vbasic.png");
            this.imageListSmall.Images.SetKeyName(21, "vcf.png");
            this.imageListSmall.Images.SetKeyName(22, "bat.png");
            // 
            // toolStripContainer1
            // 
            this.toolStripContainer1.BottomToolStripPanelVisible = false;
            // 
            // toolStripContainer1.ContentPanel
            // 
            this.toolStripContainer1.ContentPanel.Controls.Add(this.panelFileNameFilter);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.splitContainer1);
            this.toolStripContainer1.ContentPanel.Size = new System.Drawing.Size(585, 378);
            this.toolStripContainer1.ContentPanel.SizeChanged += new System.EventHandler(this.ToolStripContainer1_ContentPanel_SizeChanged);
            this.toolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStripContainer1.LeftToolStripPanelVisible = false;
            this.toolStripContainer1.Location = new System.Drawing.Point(0, 0);
            this.toolStripContainer1.Name = "toolStripContainer1";
            // 
            // toolStripContainer1.RightToolStripPanel
            // 
            this.toolStripContainer1.RightToolStripPanel.Controls.Add(this.toolStrip2);
            this.toolStripContainer1.RightToolStripPanel.Controls.Add(this.ToolStrip1);
            this.toolStripContainer1.Size = new System.Drawing.Size(609, 378);
            this.toolStripContainer1.TabIndex = 12;
            this.toolStripContainer1.Text = "toolStripContainer1";
            this.toolStripContainer1.TopToolStripPanelVisible = false;
            // 
            // panelFileNameFilter
            // 
            this.panelFileNameFilter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelFileNameFilter.Controls.Add(this.labelFileType);
            this.panelFileNameFilter.Controls.Add(this.labelFileName);
            this.panelFileNameFilter.Controls.Add(this.comboBoxFileTypes);
            this.panelFileNameFilter.Controls.Add(this.textBoxSelectedFiles);
            this.panelFileNameFilter.Location = new System.Drawing.Point(2, 324);
            this.panelFileNameFilter.Name = "panelFileNameFilter";
            this.panelFileNameFilter.Size = new System.Drawing.Size(582, 53);
            this.panelFileNameFilter.TabIndex = 11;
            // 
            // toolStrip2
            // 
            this.toolStrip2.CanOverflow = false;
            this.toolStrip2.Dock = System.Windows.Forms.DockStyle.None;
            this.toolStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StripButtonViewLargeIcon,
            this.StripButtonViewSmallIcon,
            this.StripButtonViewDetails});
            this.toolStrip2.Location = new System.Drawing.Point(0, 3);
            this.toolStrip2.Name = "toolStrip2";
            this.toolStrip2.Size = new System.Drawing.Size(24, 78);
            this.toolStrip2.TabIndex = 1;
            // 
            // StripButtonViewLargeIcon
            // 
            this.StripButtonViewLargeIcon.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonViewLargeIcon.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonViewLargeIcon.Image")));
            this.StripButtonViewLargeIcon.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonViewLargeIcon.Name = "StripButtonViewLargeIcon";
            this.StripButtonViewLargeIcon.Size = new System.Drawing.Size(22, 20);
            this.StripButtonViewLargeIcon.Text = "Large Icons";
            this.StripButtonViewLargeIcon.Click += new System.EventHandler(this.StripButtonViewLargeIcon_Click);
            // 
            // StripButtonViewSmallIcon
            // 
            this.StripButtonViewSmallIcon.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonViewSmallIcon.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonViewSmallIcon.Image")));
            this.StripButtonViewSmallIcon.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonViewSmallIcon.Name = "StripButtonViewSmallIcon";
            this.StripButtonViewSmallIcon.Size = new System.Drawing.Size(22, 20);
            this.StripButtonViewSmallIcon.Text = "Small Icons";
            this.StripButtonViewSmallIcon.Click += new System.EventHandler(this.StripButtonViewSmallIcon_Click);
            // 
            // StripButtonViewDetails
            // 
            this.StripButtonViewDetails.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonViewDetails.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonViewDetails.Image")));
            this.StripButtonViewDetails.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonViewDetails.Name = "StripButtonViewDetails";
            this.StripButtonViewDetails.Size = new System.Drawing.Size(22, 20);
            this.StripButtonViewDetails.Text = "Details";
            this.StripButtonViewDetails.Click += new System.EventHandler(this.StripButtonViewDetails_Click);
            // 
            // ToolStrip1
            // 
            this.ToolStrip1.CanOverflow = false;
            this.ToolStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.ToolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StripButtonGoUpward,
            this.StripButtonGoUndo,
            this.StripButtonGoRedo,
            this.StripButtonCreateDirectory,
            this.StripButtonDeleteDirectory,
            this.StripButtonDeleteFile});
            this.ToolStrip1.Location = new System.Drawing.Point(0, 81);
            this.ToolStrip1.Name = "ToolStrip1";
            this.ToolStrip1.Size = new System.Drawing.Size(24, 78);
            this.ToolStrip1.TabIndex = 0;
            // 
            // StripButtonGoUpward
            // 
            this.StripButtonGoUpward.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonGoUpward.Enabled = false;
            this.StripButtonGoUpward.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonGoUpward.Image")));
            this.StripButtonGoUpward.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonGoUpward.Name = "StripButtonGoUpward";
            this.StripButtonGoUpward.Size = new System.Drawing.Size(22, 20);
            this.StripButtonGoUpward.Text = "Go Upward";
            // 
            // StripButtonGoUndo
            // 
            this.StripButtonGoUndo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonGoUndo.Enabled = false;
            this.StripButtonGoUndo.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonGoUndo.Image")));
            this.StripButtonGoUndo.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonGoUndo.Name = "StripButtonGoUndo";
            this.StripButtonGoUndo.Size = new System.Drawing.Size(22, 20);
            this.StripButtonGoUndo.Text = "Go Back";
            // 
            // StripButtonGoRedo
            // 
            this.StripButtonGoRedo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonGoRedo.Enabled = false;
            this.StripButtonGoRedo.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonGoRedo.Image")));
            this.StripButtonGoRedo.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonGoRedo.Name = "StripButtonGoRedo";
            this.StripButtonGoRedo.Size = new System.Drawing.Size(22, 20);
            this.StripButtonGoRedo.Text = "Go Forward";
            // 
            // StripButtonCreateDirectory
            // 
            this.StripButtonCreateDirectory.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonCreateDirectory.Enabled = false;
            this.StripButtonCreateDirectory.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonCreateDirectory.Image")));
            this.StripButtonCreateDirectory.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonCreateDirectory.Name = "StripButtonCreateDirectory";
            this.StripButtonCreateDirectory.Size = new System.Drawing.Size(22, 20);
            this.StripButtonCreateDirectory.Text = "Add Directory";
            this.StripButtonCreateDirectory.Visible = false;
            this.StripButtonCreateDirectory.Click += new System.EventHandler(this.StripButtonCreateDirectory_Click);
            // 
            // StripButtonDeleteDirectory
            // 
            this.StripButtonDeleteDirectory.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonDeleteDirectory.Enabled = false;
            this.StripButtonDeleteDirectory.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonDeleteDirectory.Image")));
            this.StripButtonDeleteDirectory.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonDeleteDirectory.Name = "StripButtonDeleteDirectory";
            this.StripButtonDeleteDirectory.Size = new System.Drawing.Size(22, 20);
            this.StripButtonDeleteDirectory.Text = "Delete Directory";
            this.StripButtonDeleteDirectory.Visible = false;
            this.StripButtonDeleteDirectory.Click += new System.EventHandler(this.StripButtonDeleteDirectory_Click);
            // 
            // StripButtonDeleteFile
            // 
            this.StripButtonDeleteFile.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonDeleteFile.Enabled = false;
            this.StripButtonDeleteFile.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonDeleteFile.Image")));
            this.StripButtonDeleteFile.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonDeleteFile.Name = "StripButtonDeleteFile";
            this.StripButtonDeleteFile.Size = new System.Drawing.Size(22, 20);
            this.StripButtonDeleteFile.Text = "Delete File(s)";
            this.StripButtonDeleteFile.Visible = false;
            this.StripButtonDeleteFile.Click += new System.EventHandler(this.StripButtonDeleteFile_Click);
            // 
            // OpenFilePanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.toolStripContainer1);
            this.Name = "OpenFilePanel";
            this.Size = new System.Drawing.Size(609, 378);
            this.Load += new System.EventHandler(this.OpenFilePanel_Load);
            this.Resize += new System.EventHandler(this.OpenFilePanel_Resize);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.toolStripContainer1.ContentPanel.ResumeLayout(false);
            this.toolStripContainer1.RightToolStripPanel.ResumeLayout(false);
            this.toolStripContainer1.RightToolStripPanel.PerformLayout();
            this.toolStripContainer1.ResumeLayout(false);
            this.toolStripContainer1.PerformLayout();
            this.panelFileNameFilter.ResumeLayout(false);
            this.panelFileNameFilter.PerformLayout();
            this.toolStrip2.ResumeLayout(false);
            this.toolStrip2.PerformLayout();
            this.ToolStrip1.ResumeLayout(false);
            this.ToolStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label labelFileType;
        private System.Windows.Forms.Label labelFileName;
        private System.Windows.Forms.ComboBox comboBoxFileTypes;
        private System.Windows.Forms.TextBox textBoxSelectedFiles;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TreeView TreeView1;
        private System.Windows.Forms.ListView ListView1;
        private System.Windows.Forms.ColumnHeader colHeaderName;
        private System.Windows.Forms.ToolStripContainer toolStripContainer1;
        private System.Windows.Forms.ToolStrip ToolStrip1;
        private System.Windows.Forms.ToolStripButton StripButtonGoUpward;
        private System.Windows.Forms.ToolStripButton StripButtonGoUndo;
        private System.Windows.Forms.ToolStripButton StripButtonGoRedo;
        private System.Windows.Forms.ToolStripButton StripButtonCreateDirectory;
        private System.Windows.Forms.ToolStripButton StripButtonDeleteDirectory;
        private System.Windows.Forms.ToolStripButton StripButtonDeleteFile;
        private System.Windows.Forms.ImageList imageListSmall;
        private System.Windows.Forms.ImageList imageListLarge;
        private System.Windows.Forms.ImageList imageListLeft;
        private System.Windows.Forms.ToolStrip toolStrip2;
        private System.Windows.Forms.ToolStripButton StripButtonViewLargeIcon;
        private System.Windows.Forms.ToolStripButton StripButtonViewSmallIcon;
        private System.Windows.Forms.ToolStripButton StripButtonViewDetails;
        private System.Windows.Forms.ColumnHeader colHeaderSize;
        private System.Windows.Forms.Panel panelFileNameFilter;
    }
}
