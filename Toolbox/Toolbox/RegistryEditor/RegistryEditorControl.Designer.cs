namespace NetOffice.DeveloperToolbox.RegistryEditor
{
    partial class RegistryEditorControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
 
        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RegistryEditorControl));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.buttonInfo = new System.Windows.Forms.Button();
            this.labelTitle = new System.Windows.Forms.Label();
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStripKeys = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripKeyCreate = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripKeyDelete = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripKeyEdit = new System.Windows.Forms.ToolStripMenuItem();
            this.imageListRegistry = new System.Windows.Forms.ImageList(this.components);
            this.dataGridViewRegistry = new System.Windows.Forms.DataGridView();
            this.TypeIcon = new System.Windows.Forms.DataGridViewImageColumn();
            this.regName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.regType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.regValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStripEntries = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripCreateEntry = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripCreateStringEntry = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripCreateBinaryEntry = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripCreateDWORDEntry = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripEditEntryValue = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripDeleteEntry = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripEditEntryName = new System.Windows.Forms.ToolStripMenuItem();
            this.labelCurrentPath = new System.Windows.Forms.Label();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.imageListValueTypes = new System.Windows.Forms.ImageList(this.components);
            this.checkBoxDeleteQuestion = new System.Windows.Forms.CheckBox();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.labelNoAdminHint = new System.Windows.Forms.Label();
            this.labelNoAdminHintIcon = new System.Windows.Forms.Label();
            this.pictureBoxNoAdminHint = new System.Windows.Forms.PictureBox();
            this.contextMenuStripNoAdmin = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.treeViewRegistry = new NetOffice.DeveloperToolbox.MultiSelectTreeView();
            this.contextMenuStripKeys.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRegistry)).BeginInit();
            this.contextMenuStripEntries.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxNoAdminHint)).BeginInit();
            this.contextMenuStripNoAdmin.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonInfo
            // 
            this.buttonInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonInfo.Image = ((System.Drawing.Image)(resources.GetObject("buttonInfo.Image")));
            this.buttonInfo.Location = new System.Drawing.Point(877, 10);
            this.buttonInfo.Name = "buttonInfo";
            this.buttonInfo.Size = new System.Drawing.Size(28, 28);
            this.buttonInfo.TabIndex = 28;
            this.buttonInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonInfo.UseVisualStyleBackColor = true;
            this.buttonInfo.Click += new System.EventHandler(this.buttonInfo_Click);
            // 
            // labelTitle
            // 
            this.labelTitle.BackColor = System.Drawing.Color.Khaki;
            this.labelTitle.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTitle.Location = new System.Drawing.Point(40, 11);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(227, 17);
            this.labelTitle.TabIndex = 29;
            this.labelTitle.Text = "Office Registry Keys auf einen Blick.";
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRefresh.Image = ((System.Drawing.Image)(resources.GetObject("buttonRefresh.Image")));
            this.buttonRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonRefresh.Location = new System.Drawing.Point(744, 10);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new System.Drawing.Size(120, 28);
            this.buttonRefresh.TabIndex = 30;
            this.buttonRefresh.Text = "&Aktualisieren";
            this.buttonRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "Name";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.HeaderText = "Type";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn3.HeaderText = "Value";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // contextMenuStripKeys
            // 
            this.contextMenuStripKeys.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripKeyCreate,
            this.toolStripKeyDelete,
            this.toolStripSeparator1,
            this.toolStripKeyEdit});
            this.contextMenuStripKeys.Name = "contextMenuStripKeys";
            this.contextMenuStripKeys.Size = new System.Drawing.Size(149, 76);
            // 
            // toolStripKeyCreate
            // 
            this.toolStripKeyCreate.Image = ((System.Drawing.Image)(resources.GetObject("toolStripKeyCreate.Image")));
            this.toolStripKeyCreate.Name = "toolStripKeyCreate";
            this.toolStripKeyCreate.Size = new System.Drawing.Size(148, 22);
            this.toolStripKeyCreate.Text = "Neu";
            this.toolStripKeyCreate.Click += new System.EventHandler(this.toolStripKeyCreate_Click);
            // 
            // toolStripKeyDelete
            // 
            this.toolStripKeyDelete.Image = ((System.Drawing.Image)(resources.GetObject("toolStripKeyDelete.Image")));
            this.toolStripKeyDelete.Name = "toolStripKeyDelete";
            this.toolStripKeyDelete.Size = new System.Drawing.Size(148, 22);
            this.toolStripKeyDelete.Text = "Löschen";
            this.toolStripKeyDelete.Click += new System.EventHandler(this.toolStripKeyDelete_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(145, 6);
            // 
            // toolStripKeyEdit
            // 
            this.toolStripKeyEdit.Enabled = false;
            this.toolStripKeyEdit.Image = ((System.Drawing.Image)(resources.GetObject("toolStripKeyEdit.Image")));
            this.toolStripKeyEdit.Name = "toolStripKeyEdit";
            this.toolStripKeyEdit.Size = new System.Drawing.Size(148, 22);
            this.toolStripKeyEdit.Text = "Umbenennen";
            this.toolStripKeyEdit.Click += new System.EventHandler(this.toolStripKeyEdit_Click);
            // 
            // imageListRegistry
            // 
            this.imageListRegistry.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListRegistry.ImageStream")));
            this.imageListRegistry.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListRegistry.Images.SetKeyName(0, "CLSDFOLD.ICO");
            this.imageListRegistry.Images.SetKeyName(1, "OPENFOLD.ICO");
            // 
            // dataGridViewRegistry
            // 
            this.dataGridViewRegistry.AllowUserToAddRows = false;
            this.dataGridViewRegistry.AllowUserToDeleteRows = false;
            this.dataGridViewRegistry.AllowUserToResizeRows = false;
            this.dataGridViewRegistry.BackgroundColor = System.Drawing.SystemColors.ActiveCaptionText;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewRegistry.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewRegistry.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewRegistry.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.TypeIcon,
            this.regName,
            this.regType,
            this.regValue});
            this.dataGridViewRegistry.ContextMenuStrip = this.contextMenuStripEntries;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewRegistry.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewRegistry.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewRegistry.GridColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.dataGridViewRegistry.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewRegistry.Name = "dataGridViewRegistry";
            this.dataGridViewRegistry.RowHeadersVisible = false;
            this.dataGridViewRegistry.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewRegistry.Size = new System.Drawing.Size(606, 423);
            this.dataGridViewRegistry.TabIndex = 31;
            this.dataGridViewRegistry.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewRegistry_CellDoubleClick);
            this.dataGridViewRegistry.SelectionChanged += new System.EventHandler(this.dataGridViewRegistry_SelectionChanged);
            this.dataGridViewRegistry.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridViewRegistry_KeyDown);
            // 
            // TypeIcon
            // 
            this.TypeIcon.Frozen = true;
            this.TypeIcon.HeaderText = "";
            this.TypeIcon.Name = "TypeIcon";
            this.TypeIcon.ReadOnly = true;
            this.TypeIcon.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.TypeIcon.Width = 30;
            // 
            // regName
            // 
            this.regName.Frozen = true;
            this.regName.HeaderText = "Name";
            this.regName.Name = "regName";
            this.regName.ReadOnly = true;
            this.regName.Width = 180;
            // 
            // regType
            // 
            this.regType.HeaderText = "Type";
            this.regType.Name = "regType";
            this.regType.ReadOnly = true;
            // 
            // regValue
            // 
            this.regValue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.regValue.HeaderText = "Value";
            this.regValue.Name = "regValue";
            this.regValue.ReadOnly = true;
            // 
            // contextMenuStripEntries
            // 
            this.contextMenuStripEntries.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripCreateEntry,
            this.toolStripSeparator2,
            this.toolStripEditEntryValue,
            this.toolStripDeleteEntry,
            this.toolStripEditEntryName});
            this.contextMenuStripEntries.Name = "contextMenuStripEntries";
            this.contextMenuStripEntries.Size = new System.Drawing.Size(149, 98);
            this.contextMenuStripEntries.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuStripEntries_Opening);
            // 
            // toolStripCreateEntry
            // 
            this.toolStripCreateEntry.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripCreateStringEntry,
            this.toolStripCreateBinaryEntry,
            this.toolStripCreateDWORDEntry});
            this.toolStripCreateEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripCreateEntry.Image")));
            this.toolStripCreateEntry.Name = "toolStripCreateEntry";
            this.toolStripCreateEntry.Size = new System.Drawing.Size(148, 22);
            this.toolStripCreateEntry.Text = "Neuer Wert";
            // 
            // toolStripCreateStringEntry
            // 
            this.toolStripCreateStringEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripCreateStringEntry.Image")));
            this.toolStripCreateStringEntry.Name = "toolStripCreateStringEntry";
            this.toolStripCreateStringEntry.Size = new System.Drawing.Size(146, 22);
            this.toolStripCreateStringEntry.Text = "Zeichenfolge";
            this.toolStripCreateStringEntry.Click += new System.EventHandler(this.toolStripCreateStringEntry_Click);
            // 
            // toolStripCreateBinaryEntry
            // 
            this.toolStripCreateBinaryEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripCreateBinaryEntry.Image")));
            this.toolStripCreateBinaryEntry.Name = "toolStripCreateBinaryEntry";
            this.toolStripCreateBinaryEntry.Size = new System.Drawing.Size(146, 22);
            this.toolStripCreateBinaryEntry.Text = "Binärwert";
            this.toolStripCreateBinaryEntry.Click += new System.EventHandler(this.toolStripCreateBinaryEntry_Click);
            // 
            // toolStripCreateDWORDEntry
            // 
            this.toolStripCreateDWORDEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripCreateDWORDEntry.Image")));
            this.toolStripCreateDWORDEntry.Name = "toolStripCreateDWORDEntry";
            this.toolStripCreateDWORDEntry.Size = new System.Drawing.Size(146, 22);
            this.toolStripCreateDWORDEntry.Text = "DWORD";
            this.toolStripCreateDWORDEntry.Click += new System.EventHandler(this.toolStripCreateDWORDEntry_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(145, 6);
            // 
            // toolStripEditEntryValue
            // 
            this.toolStripEditEntryValue.Image = ((System.Drawing.Image)(resources.GetObject("toolStripEditEntryValue.Image")));
            this.toolStripEditEntryValue.Name = "toolStripEditEntryValue";
            this.toolStripEditEntryValue.Size = new System.Drawing.Size(148, 22);
            this.toolStripEditEntryValue.Text = "Wert ändern";
            this.toolStripEditEntryValue.Click += new System.EventHandler(this.toolStripEditEntryValue_Click);
            // 
            // toolStripDeleteEntry
            // 
            this.toolStripDeleteEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDeleteEntry.Image")));
            this.toolStripDeleteEntry.Name = "toolStripDeleteEntry";
            this.toolStripDeleteEntry.Size = new System.Drawing.Size(148, 22);
            this.toolStripDeleteEntry.Text = "Löschen";
            this.toolStripDeleteEntry.Click += new System.EventHandler(this.toolStripDeleteEntry_Click);
            // 
            // toolStripEditEntryName
            // 
            this.toolStripEditEntryName.Image = ((System.Drawing.Image)(resources.GetObject("toolStripEditEntryName.Image")));
            this.toolStripEditEntryName.Name = "toolStripEditEntryName";
            this.toolStripEditEntryName.Size = new System.Drawing.Size(148, 22);
            this.toolStripEditEntryName.Text = "Umbenennen";
            this.toolStripEditEntryName.Click += new System.EventHandler(this.toolStripEditEntryName_Click);
            // 
            // labelCurrentPath
            // 
            this.labelCurrentPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelCurrentPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelCurrentPath.Location = new System.Drawing.Point(3, 480);
            this.labelCurrentPath.Name = "labelCurrentPath";
            this.labelCurrentPath.Size = new System.Drawing.Size(917, 15);
            this.labelCurrentPath.TabIndex = 33;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.splitContainer1.Location = new System.Drawing.Point(3, 55);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.treeViewRegistry);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridViewRegistry);
            this.splitContainer1.Size = new System.Drawing.Size(918, 427);
            this.splitContainer1.SplitterDistance = 304;
            this.splitContainer1.TabIndex = 34;
            // 
            // imageListValueTypes
            // 
            this.imageListValueTypes.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListValueTypes.ImageStream")));
            this.imageListValueTypes.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListValueTypes.Images.SetKeyName(0, "StringType.ico");
            this.imageListValueTypes.Images.SetKeyName(1, "OtherType.ico");
            // 
            // checkBoxDeleteQuestion
            // 
            this.checkBoxDeleteQuestion.AutoSize = true;
            this.checkBoxDeleteQuestion.Checked = true;
            this.checkBoxDeleteQuestion.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxDeleteQuestion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxDeleteQuestion.Location = new System.Drawing.Point(45, 33);
            this.checkBoxDeleteQuestion.Name = "checkBoxDeleteQuestion";
            this.checkBoxDeleteQuestion.Size = new System.Drawing.Size(202, 20);
            this.checkBoxDeleteQuestion.TabIndex = 36;
            this.checkBoxDeleteQuestion.Text = "Vor dem Löschen nachfragen";
            this.checkBoxDeleteQuestion.UseVisualStyleBackColor = true;
            // 
            // pictureBox4
            // 
            this.pictureBox4.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox4.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox4.Image")));
            this.pictureBox4.Location = new System.Drawing.Point(19, 10);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(16, 16);
            this.pictureBox4.TabIndex = 40;
            this.pictureBox4.TabStop = false;
            // 
            // labelNoAdminHint
            // 
            this.labelNoAdminHint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelNoAdminHint.BackColor = System.Drawing.Color.Khaki;
            this.labelNoAdminHint.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelNoAdminHint.Location = new System.Drawing.Point(415, 10);
            this.labelNoAdminHint.Name = "labelNoAdminHint";
            this.labelNoAdminHint.Size = new System.Drawing.Size(323, 28);
            this.labelNoAdminHint.TabIndex = 41;
            this.labelNoAdminHint.Text = "Aufgrund fehlender Administratorberechtigung können Sie keine Werte im HiveKey Lo" +
                "calMachine hinzufügen, löschen oder ändern.";
            this.labelNoAdminHint.Visible = false;
            // 
            // labelNoAdminHintIcon
            // 
            this.labelNoAdminHintIcon.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelNoAdminHintIcon.BackColor = System.Drawing.Color.Khaki;
            this.labelNoAdminHintIcon.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelNoAdminHintIcon.Location = new System.Drawing.Point(397, 10);
            this.labelNoAdminHintIcon.Name = "labelNoAdminHintIcon";
            this.labelNoAdminHintIcon.Size = new System.Drawing.Size(18, 28);
            this.labelNoAdminHintIcon.TabIndex = 73;
            this.labelNoAdminHintIcon.Visible = false;
            // 
            // pictureBoxNoAdminHint
            // 
            this.pictureBoxNoAdminHint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBoxNoAdminHint.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxNoAdminHint.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxNoAdminHint.Image")));
            this.pictureBoxNoAdminHint.Location = new System.Drawing.Point(396, 10);
            this.pictureBoxNoAdminHint.Name = "pictureBoxNoAdminHint";
            this.pictureBoxNoAdminHint.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxNoAdminHint.TabIndex = 74;
            this.pictureBoxNoAdminHint.TabStop = false;
            this.pictureBoxNoAdminHint.Visible = false;
            // 
            // contextMenuStripNoAdmin
            // 
            this.contextMenuStripNoAdmin.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.contextMenuStripNoAdmin.Name = "contextMenuStripNoAdmin";
            this.contextMenuStripNoAdmin.Size = new System.Drawing.Size(216, 26);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Enabled = false;
            this.toolStripMenuItem1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripMenuItem1.Image")));
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(215, 22);
            this.toolStripMenuItem1.Text = "Keine Administrator Rechte";
            // 
            // treeViewRegistry
            // 
            this.treeViewRegistry.ContextMenuStrip = this.contextMenuStripEntries;
            this.treeViewRegistry.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewRegistry.DrawMode = System.Windows.Forms.TreeViewDrawMode.OwnerDrawText;
            this.treeViewRegistry.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewRegistry.HideSelection = false;
            this.treeViewRegistry.ImageIndex = 0;
            this.treeViewRegistry.ImageList = this.imageListRegistry;
            this.treeViewRegistry.Location = new System.Drawing.Point(0, 0);
            this.treeViewRegistry.Name = "treeViewRegistry";
            this.treeViewRegistry.SelectedImageIndex = 1;
            this.treeViewRegistry.SelectedNodes = ((System.Collections.Generic.List<System.Windows.Forms.TreeNode>)(resources.GetObject("treeViewRegistry.SelectedNodes")));
            this.treeViewRegistry.Size = new System.Drawing.Size(300, 423);
            this.treeViewRegistry.TabIndex = 33;
            this.treeViewRegistry.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.treeViewRegistry_AfterLabelEdit);
            this.treeViewRegistry.AfterCollapse += new System.Windows.Forms.TreeViewEventHandler(this.treeViewRegistry_AfterCollapse);
            this.treeViewRegistry.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeViewRegistry_BeforeExpand);
            this.treeViewRegistry.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.treeViewRegistry_AfterExpand);
            this.treeViewRegistry.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeViewRegistry_AfterSelect);
            this.treeViewRegistry.KeyDown += new System.Windows.Forms.KeyEventHandler(this.treeViewRegistry_KeyDown);
            // 
            // RegistryEditorControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pictureBoxNoAdminHint);
            this.Controls.Add(this.labelNoAdminHintIcon);
            this.Controls.Add(this.labelNoAdminHint);
            this.Controls.Add(this.pictureBox4);
            this.Controls.Add(this.checkBoxDeleteQuestion);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.labelCurrentPath);
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.labelTitle);
            this.Controls.Add(this.buttonInfo);
            this.Name = "RegistryEditorControl";
            this.Size = new System.Drawing.Size(924, 496);
            this.contextMenuStripKeys.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRegistry)).EndInit();
            this.contextMenuStripEntries.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxNoAdminHint)).EndInit();
            this.contextMenuStripNoAdmin.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        private System.Windows.Forms.Button buttonInfo;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Button buttonRefresh;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridView dataGridViewRegistry;
        private System.Windows.Forms.Label labelCurrentPath;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripKeys;
        private System.Windows.Forms.ToolStripMenuItem toolStripKeyCreate;
        private System.Windows.Forms.ToolStripMenuItem toolStripKeyDelete;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripEntries;
        private System.Windows.Forms.ToolStripMenuItem toolStripCreateEntry;
        private System.Windows.Forms.ToolStripMenuItem toolStripCreateStringEntry;
        private System.Windows.Forms.ToolStripMenuItem toolStripCreateBinaryEntry;
        private System.Windows.Forms.ToolStripMenuItem toolStripCreateDWORDEntry;
        private System.Windows.Forms.ToolStripMenuItem toolStripDeleteEntry;
        private System.Windows.Forms.ImageList imageListRegistry;
        private System.Windows.Forms.ImageList imageListValueTypes;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripKeyEdit;
        private System.Windows.Forms.ToolStripMenuItem toolStripEditEntryValue;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem toolStripEditEntryName;
        private System.Windows.Forms.CheckBox checkBoxDeleteQuestion;
        private System.Windows.Forms.DataGridViewImageColumn TypeIcon;
        private System.Windows.Forms.DataGridViewTextBoxColumn regName;
        private System.Windows.Forms.DataGridViewTextBoxColumn regType;
        private System.Windows.Forms.DataGridViewTextBoxColumn regValue;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.Label labelNoAdminHint;
        private System.Windows.Forms.Label labelNoAdminHintIcon;
        private System.Windows.Forms.PictureBox pictureBoxNoAdminHint;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripNoAdmin;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private MultiSelectTreeView treeViewRegistry;
    }
}
