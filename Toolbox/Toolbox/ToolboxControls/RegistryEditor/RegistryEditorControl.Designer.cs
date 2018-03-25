namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.labelTitle = new System.Windows.Forms.Label();
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.contextMenuStripKeys = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripKeyCreate = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripKeyDelete = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripKeyEdit = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripKeyExport = new System.Windows.Forms.ToolStripMenuItem();
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
            this.labelNoAdminHint = new System.Windows.Forms.Label();
            this.contextMenuStripNoAdmin = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.labelSearch = new System.Windows.Forms.Label();
            this.pictureBoxNoResult = new System.Windows.Forms.PictureBox();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SearchBoxPanel = new System.Windows.Forms.Panel();
            this.SearchLabel = new System.Windows.Forms.Label();
            this.textBoxSearch = new System.Windows.Forms.TextBox();
            this.pictureBoxSearching = new System.Windows.Forms.PictureBox();
            this.treeViewRegistry = new NetOffice.DeveloperToolbox.Controls.Tree.MultiSelectTreeView();
            this.contextMenuStripKeys.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRegistry)).BeginInit();
            this.contextMenuStripEntries.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.contextMenuStripNoAdmin.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxNoResult)).BeginInit();
            this.SearchBoxPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSearching)).BeginInit();
            this.SuspendLayout();
            // 
            // labelTitle
            // 
            this.labelTitle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelTitle.BackColor = System.Drawing.Color.Orange;
            this.labelTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelTitle.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelTitle.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTitle.ForeColor = System.Drawing.Color.Black;
            this.labelTitle.Image = ((System.Drawing.Image)(resources.GetObject("labelTitle.Image")));
            this.labelTitle.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.labelTitle.Location = new System.Drawing.Point(3, 1);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(910, 28);
            this.labelTitle.TabIndex = 29;
            this.labelTitle.Text = "      Office Registry Keys at a glance";
            this.labelTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonRefresh.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonRefresh.FlatAppearance.BorderSize = 0;
            this.buttonRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRefresh.ForeColor = System.Drawing.Color.Blue;
            this.buttonRefresh.Image = ((System.Drawing.Image)(resources.GetObject("buttonRefresh.Image")));
            this.buttonRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonRefresh.Location = new System.Drawing.Point(792, 611);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new System.Drawing.Size(120, 32);
            this.buttonRefresh.TabIndex = 30;
            this.buttonRefresh.Text = "&Refresh";
            this.buttonRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // contextMenuStripKeys
            // 
            this.contextMenuStripKeys.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripKeyCreate,
            this.toolStripKeyDelete,
            this.toolStripSeparator1,
            this.toolStripKeyEdit,
            this.toolStripMenuItem2,
            this.toolStripKeyExport});
            this.contextMenuStripKeys.Name = "contextMenuStripKeys";
            this.contextMenuStripKeys.Size = new System.Drawing.Size(118, 104);
            // 
            // toolStripKeyCreate
            // 
            this.toolStripKeyCreate.Image = ((System.Drawing.Image)(resources.GetObject("toolStripKeyCreate.Image")));
            this.toolStripKeyCreate.Name = "toolStripKeyCreate";
            this.toolStripKeyCreate.Size = new System.Drawing.Size(117, 22);
            this.toolStripKeyCreate.Text = "New";
            this.toolStripKeyCreate.Click += new System.EventHandler(this.toolStripKeyCreate_Click);
            // 
            // toolStripKeyDelete
            // 
            this.toolStripKeyDelete.Image = ((System.Drawing.Image)(resources.GetObject("toolStripKeyDelete.Image")));
            this.toolStripKeyDelete.Name = "toolStripKeyDelete";
            this.toolStripKeyDelete.Size = new System.Drawing.Size(117, 22);
            this.toolStripKeyDelete.Text = "Delete";
            this.toolStripKeyDelete.Click += new System.EventHandler(this.toolStripKeyDelete_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(114, 6);
            // 
            // toolStripKeyEdit
            // 
            this.toolStripKeyEdit.Enabled = false;
            this.toolStripKeyEdit.Image = ((System.Drawing.Image)(resources.GetObject("toolStripKeyEdit.Image")));
            this.toolStripKeyEdit.Name = "toolStripKeyEdit";
            this.toolStripKeyEdit.Size = new System.Drawing.Size(117, 22);
            this.toolStripKeyEdit.Text = "Rename";
            this.toolStripKeyEdit.Click += new System.EventHandler(this.toolStripKeyEdit_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(114, 6);
            // 
            // toolStripKeyExport
            // 
            this.toolStripKeyExport.Image = ((System.Drawing.Image)(resources.GetObject("toolStripKeyExport.Image")));
            this.toolStripKeyExport.Name = "toolStripKeyExport";
            this.toolStripKeyExport.Size = new System.Drawing.Size(117, 22);
            this.toolStripKeyExport.Text = "Export";
            this.toolStripKeyExport.Click += new System.EventHandler(this.toolStripKeyExport_Click);
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
            this.dataGridViewRegistry.BackgroundColor = System.Drawing.Color.White;
            this.dataGridViewRegistry.BorderStyle = System.Windows.Forms.BorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.LightSteelBlue;
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
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewRegistry.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewRegistry.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewRegistry.GridColor = System.Drawing.Color.White;
            this.dataGridViewRegistry.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewRegistry.Name = "dataGridViewRegistry";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewRegistry.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewRegistry.RowHeadersVisible = false;
            this.dataGridViewRegistry.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewRegistry.ShowCellErrors = false;
            this.dataGridViewRegistry.ShowRowErrors = false;
            this.dataGridViewRegistry.Size = new System.Drawing.Size(603, 582);
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
            this.contextMenuStripEntries.Size = new System.Drawing.Size(147, 98);
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
            this.toolStripCreateEntry.Size = new System.Drawing.Size(146, 22);
            this.toolStripCreateEntry.Text = "New Value";
            // 
            // toolStripCreateStringEntry
            // 
            this.toolStripCreateStringEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripCreateStringEntry.Image")));
            this.toolStripCreateStringEntry.Name = "toolStripCreateStringEntry";
            this.toolStripCreateStringEntry.Size = new System.Drawing.Size(117, 22);
            this.toolStripCreateStringEntry.Text = "String";
            this.toolStripCreateStringEntry.Click += new System.EventHandler(this.toolStripCreateStringEntry_Click);
            // 
            // toolStripCreateBinaryEntry
            // 
            this.toolStripCreateBinaryEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripCreateBinaryEntry.Image")));
            this.toolStripCreateBinaryEntry.Name = "toolStripCreateBinaryEntry";
            this.toolStripCreateBinaryEntry.Size = new System.Drawing.Size(117, 22);
            this.toolStripCreateBinaryEntry.Text = "Binary";
            this.toolStripCreateBinaryEntry.Click += new System.EventHandler(this.toolStripCreateBinaryEntry_Click);
            // 
            // toolStripCreateDWORDEntry
            // 
            this.toolStripCreateDWORDEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripCreateDWORDEntry.Image")));
            this.toolStripCreateDWORDEntry.Name = "toolStripCreateDWORDEntry";
            this.toolStripCreateDWORDEntry.Size = new System.Drawing.Size(117, 22);
            this.toolStripCreateDWORDEntry.Text = "DWORD";
            this.toolStripCreateDWORDEntry.Click += new System.EventHandler(this.toolStripCreateDWORDEntry_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(143, 6);
            // 
            // toolStripEditEntryValue
            // 
            this.toolStripEditEntryValue.Image = ((System.Drawing.Image)(resources.GetObject("toolStripEditEntryValue.Image")));
            this.toolStripEditEntryValue.Name = "toolStripEditEntryValue";
            this.toolStripEditEntryValue.Size = new System.Drawing.Size(146, 22);
            this.toolStripEditEntryValue.Text = "Change Value";
            this.toolStripEditEntryValue.Click += new System.EventHandler(this.toolStripEditEntryValue_Click);
            // 
            // toolStripDeleteEntry
            // 
            this.toolStripDeleteEntry.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDeleteEntry.Image")));
            this.toolStripDeleteEntry.Name = "toolStripDeleteEntry";
            this.toolStripDeleteEntry.Size = new System.Drawing.Size(146, 22);
            this.toolStripDeleteEntry.Text = "Delete";
            this.toolStripDeleteEntry.Click += new System.EventHandler(this.toolStripDeleteEntry_Click);
            // 
            // toolStripEditEntryName
            // 
            this.toolStripEditEntryName.Image = ((System.Drawing.Image)(resources.GetObject("toolStripEditEntryName.Image")));
            this.toolStripEditEntryName.Name = "toolStripEditEntryName";
            this.toolStripEditEntryName.Size = new System.Drawing.Size(146, 22);
            this.toolStripEditEntryName.Text = "Rename";
            this.toolStripEditEntryName.Click += new System.EventHandler(this.toolStripEditEntryName_Click);
            // 
            // labelCurrentPath
            // 
            this.labelCurrentPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelCurrentPath.AutoEllipsis = true;
            this.labelCurrentPath.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCurrentPath.ForeColor = System.Drawing.Color.Black;
            this.labelCurrentPath.Location = new System.Drawing.Point(3, 618);
            this.labelCurrentPath.Name = "labelCurrentPath";
            this.labelCurrentPath.Size = new System.Drawing.Size(783, 31);
            this.labelCurrentPath.TabIndex = 33;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitContainer1.Location = new System.Drawing.Point(3, 27);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.treeViewRegistry);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridViewRegistry);
            this.splitContainer1.Size = new System.Drawing.Size(910, 584);
            this.splitContainer1.SplitterDistance = 301;
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
            this.checkBoxDeleteQuestion.BackColor = System.Drawing.Color.Orange;
            this.checkBoxDeleteQuestion.Checked = true;
            this.checkBoxDeleteQuestion.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxDeleteQuestion.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxDeleteQuestion.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxDeleteQuestion.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxDeleteQuestion.Location = new System.Drawing.Point(303, 5);
            this.checkBoxDeleteQuestion.Name = "checkBoxDeleteQuestion";
            this.checkBoxDeleteQuestion.Size = new System.Drawing.Size(138, 21);
            this.checkBoxDeleteQuestion.TabIndex = 36;
            this.checkBoxDeleteQuestion.Text = "Ask before deleting";
            this.checkBoxDeleteQuestion.UseVisualStyleBackColor = false;
            // 
            // labelNoAdminHint
            // 
            this.labelNoAdminHint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelNoAdminHint.BackColor = System.Drawing.Color.LightGray;
            this.labelNoAdminHint.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelNoAdminHint.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelNoAdminHint.ForeColor = System.Drawing.Color.Black;
            this.labelNoAdminHint.Location = new System.Drawing.Point(32, 619);
            this.labelNoAdminHint.Name = "labelNoAdminHint";
            this.labelNoAdminHint.Size = new System.Drawing.Size(712, 16);
            this.labelNoAdminHint.TabIndex = 41;
            this.labelNoAdminHint.Text = "Due to missing administrator privileges you cannot add, delete or change values i" +
    "n the HiveKey LocalMachine.";
            this.labelNoAdminHint.Visible = false;
            // 
            // contextMenuStripNoAdmin
            // 
            this.contextMenuStripNoAdmin.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.contextMenuStripNoAdmin.Name = "contextMenuStripNoAdmin";
            this.contextMenuStripNoAdmin.Size = new System.Drawing.Size(220, 26);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Enabled = false;
            this.toolStripMenuItem1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripMenuItem1.Image")));
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(219, 22);
            this.toolStripMenuItem1.Text = "No Administrator Privileges";
            // 
            // labelSearch
            // 
            this.labelSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelSearch.AutoSize = true;
            this.labelSearch.BackColor = System.Drawing.Color.Orange;
            this.labelSearch.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSearch.ForeColor = System.Drawing.Color.Blue;
            this.labelSearch.Location = new System.Drawing.Point(647, 6);
            this.labelSearch.Name = "labelSearch";
            this.labelSearch.Size = new System.Drawing.Size(47, 17);
            this.labelSearch.TabIndex = 43;
            this.labelSearch.Text = "Search";
            this.labelSearch.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // pictureBoxNoResult
            // 
            this.pictureBoxNoResult.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBoxNoResult.BackColor = System.Drawing.Color.Orange;
            this.pictureBoxNoResult.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxNoResult.Image")));
            this.pictureBoxNoResult.Location = new System.Drawing.Point(691, 7);
            this.pictureBoxNoResult.Name = "pictureBoxNoResult";
            this.pictureBoxNoResult.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxNoResult.TabIndex = 82;
            this.pictureBoxNoResult.TabStop = false;
            this.pictureBoxNoResult.Visible = false;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.Frozen = true;
            this.dataGridViewTextBoxColumn1.HeaderText = "Name";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 180;
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
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            // 
            // SearchBoxPanel
            // 
            this.SearchBoxPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SearchBoxPanel.BackColor = System.Drawing.Color.White;
            this.SearchBoxPanel.Controls.Add(this.SearchLabel);
            this.SearchBoxPanel.Controls.Add(this.textBoxSearch);
            this.SearchBoxPanel.ForeColor = System.Drawing.Color.Black;
            this.SearchBoxPanel.Location = new System.Drawing.Point(709, 6);
            this.SearchBoxPanel.Margin = new System.Windows.Forms.Padding(0);
            this.SearchBoxPanel.Name = "SearchBoxPanel";
            this.SearchBoxPanel.Size = new System.Drawing.Size(190, 18);
            this.SearchBoxPanel.TabIndex = 83;
            // 
            // SearchLabel
            // 
            this.SearchLabel.Image = ((System.Drawing.Image)(resources.GetObject("SearchLabel.Image")));
            this.SearchLabel.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.SearchLabel.Location = new System.Drawing.Point(166, -1);
            this.SearchLabel.Name = "SearchLabel";
            this.SearchLabel.Size = new System.Drawing.Size(24, 19);
            this.SearchLabel.TabIndex = 32;
            this.SearchLabel.Click += new System.EventHandler(this.SearchLabel_Click);
            // 
            // textBoxSearch
            // 
            this.textBoxSearch.BackColor = System.Drawing.Color.White;
            this.textBoxSearch.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBoxSearch.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxSearch.ForeColor = System.Drawing.Color.Black;
            this.textBoxSearch.Location = new System.Drawing.Point(5, 0);
            this.textBoxSearch.Margin = new System.Windows.Forms.Padding(0);
            this.textBoxSearch.Name = "textBoxSearch";
            this.textBoxSearch.Size = new System.Drawing.Size(166, 18);
            this.textBoxSearch.TabIndex = 42;
            this.textBoxSearch.TextChanged += new System.EventHandler(this.textBoxSearch_TextChanged);
            this.textBoxSearch.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxSearch_KeyDown);
            this.textBoxSearch.Leave += new System.EventHandler(this.textBoxSearch_Leave);
            // 
            // pictureBoxSearching
            // 
            this.pictureBoxSearching.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBoxSearching.BackColor = System.Drawing.Color.Orange;
            this.pictureBoxSearching.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxSearching.Image")));
            this.pictureBoxSearching.Location = new System.Drawing.Point(691, 7);
            this.pictureBoxSearching.Name = "pictureBoxSearching";
            this.pictureBoxSearching.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxSearching.TabIndex = 84;
            this.pictureBoxSearching.TabStop = false;
            this.pictureBoxSearching.Visible = false;
            // 
            // treeViewRegistry
            // 
            this.treeViewRegistry.BackColor = System.Drawing.Color.White;
            this.treeViewRegistry.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.treeViewRegistry.ContextMenuStrip = this.contextMenuStripEntries;
            this.treeViewRegistry.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewRegistry.DrawMode = System.Windows.Forms.TreeViewDrawMode.OwnerDrawText;
            this.treeViewRegistry.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewRegistry.HideSelection = false;
            this.treeViewRegistry.ImageIndex = 0;
            this.treeViewRegistry.ImageList = this.imageListRegistry;
            this.treeViewRegistry.Location = new System.Drawing.Point(0, 0);
            this.treeViewRegistry.Name = "treeViewRegistry";
            this.treeViewRegistry.SelectedImageIndex = 1;
            this.treeViewRegistry.SelectedNodes = ((System.Collections.Generic.List<System.Windows.Forms.TreeNode>)(resources.GetObject("treeViewRegistry.SelectedNodes")));
            this.treeViewRegistry.Size = new System.Drawing.Size(299, 582);
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
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.pictureBoxSearching);
            this.Controls.Add(this.pictureBoxNoResult);
            this.Controls.Add(this.SearchBoxPanel);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.labelNoAdminHint);
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.labelSearch);
            this.Controls.Add(this.checkBoxDeleteQuestion);
            this.Controls.Add(this.labelCurrentPath);
            this.Controls.Add(this.labelTitle);
            this.Name = "RegistryEditorControl";
            this.Size = new System.Drawing.Size(916, 641);
            this.Resize += new System.EventHandler(this.RegistryEditorControl_Resize);
            this.contextMenuStripKeys.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRegistry)).EndInit();
            this.contextMenuStripEntries.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.contextMenuStripNoAdmin.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxNoResult)).EndInit();
            this.SearchBoxPanel.ResumeLayout(false);
            this.SearchBoxPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSearching)).EndInit();
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
        private System.Windows.Forms.Label labelNoAdminHint;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripNoAdmin;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private NetOffice.DeveloperToolbox.Controls.Tree.MultiSelectTreeView treeViewRegistry;
        private System.Windows.Forms.Label labelSearch;
        private System.Windows.Forms.PictureBox pictureBoxNoResult;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem toolStripKeyExport;
        private System.Windows.Forms.Panel SearchBoxPanel;
        private System.Windows.Forms.TextBox textBoxSearch;
        private System.Windows.Forms.PictureBox pictureBoxSearching;
        private System.Windows.Forms.Label SearchLabel;
    }
}
