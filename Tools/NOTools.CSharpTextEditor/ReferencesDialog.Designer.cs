namespace NOTools.CSharpTextEditor
{
    partial class ReferencesDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem("Please wait...");
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageGac = new System.Windows.Forms.TabPage();
            this.labelNameFilter = new System.Windows.Forms.Label();
            this.textBoxNameFilter = new System.Windows.Forms.TextBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tabPageFileSystem = new System.Windows.Forms.TabPage();
            this.openFilePanel1 = new NOTools.FileSystemDialogs.OpenFilePanel();
            this.buttonOk = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPageGac.SuspendLayout();
            this.tabPageFileSystem.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPageGac);
            this.tabControl1.Controls.Add(this.tabPageFileSystem);
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(694, 423);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPageGac
            // 
            this.tabPageGac.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tabPageGac.Controls.Add(this.labelNameFilter);
            this.tabPageGac.Controls.Add(this.textBoxNameFilter);
            this.tabPageGac.Controls.Add(this.listView1);
            this.tabPageGac.Location = new System.Drawing.Point(4, 25);
            this.tabPageGac.Margin = new System.Windows.Forms.Padding(4);
            this.tabPageGac.Name = "tabPageGac";
            this.tabPageGac.Padding = new System.Windows.Forms.Padding(4);
            this.tabPageGac.Size = new System.Drawing.Size(686, 394);
            this.tabPageGac.TabIndex = 0;
            this.tabPageGac.Text = "GAC";
            // 
            // labelNameFilter
            // 
            this.labelNameFilter.AutoSize = true;
            this.labelNameFilter.Location = new System.Drawing.Point(17, 18);
            this.labelNameFilter.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelNameFilter.Name = "labelNameFilter";
            this.labelNameFilter.Size = new System.Drawing.Size(80, 16);
            this.labelNameFilter.TabIndex = 2;
            this.labelNameFilter.Text = "Name Filter:";
            // 
            // textBoxNameFilter
            // 
            this.textBoxNameFilter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxNameFilter.Location = new System.Drawing.Point(107, 14);
            this.textBoxNameFilter.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxNameFilter.Name = "textBoxNameFilter";
            this.textBoxNameFilter.Size = new System.Drawing.Size(556, 22);
            this.textBoxNameFilter.TabIndex = 1;
            this.textBoxNameFilter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxNameFilter_KeyDown);
            // 
            // listView1
            // 
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listView1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader4});
            this.listView1.FullRowSelect = true;
            this.listView1.GridLines = true;
            this.listView1.HideSelection = false;
            this.listView1.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem3});
            this.listView1.LabelWrap = false;
            this.listView1.Location = new System.Drawing.Point(0, 47);
            this.listView1.Margin = new System.Windows.Forms.Padding(4);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(685, 337);
            this.listView1.TabIndex = 2;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.listView1_ColumnClick);
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            this.listView1.DoubleClick += new System.EventHandler(this.listView1_DoubleClick);
            this.listView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.listView1_KeyDown);
            this.listView1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listView1_MouseDoubleClick);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Name";
            this.columnHeader1.Width = 250;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Version";
            this.columnHeader2.Width = 80;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Public Key Token";
            this.columnHeader3.Width = 130;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Path";
            this.columnHeader4.Width = 205;
            // 
            // tabPageFileSystem
            // 
            this.tabPageFileSystem.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tabPageFileSystem.Controls.Add(this.openFilePanel1);
            this.tabPageFileSystem.Location = new System.Drawing.Point(4, 25);
            this.tabPageFileSystem.Margin = new System.Windows.Forms.Padding(4);
            this.tabPageFileSystem.Name = "tabPageFileSystem";
            this.tabPageFileSystem.Padding = new System.Windows.Forms.Padding(4);
            this.tabPageFileSystem.Size = new System.Drawing.Size(686, 394);
            this.tabPageFileSystem.TabIndex = 1;
            this.tabPageFileSystem.Text = "File System";
            // 
            // openFilePanel1
            // 
            this.openFilePanel1.Default.AllowAddFolders = false;
            this.openFilePanel1.Default.AllowBrowseFolders = true;
            this.openFilePanel1.Default.AllowDeleteFiles = false;
            this.openFilePanel1.Default.AllowDeleteFolders = false;
            this.openFilePanel1.Default.AllowMultipleSelect = true;
            this.openFilePanel1.Default.Visible = true;
            this.openFilePanel1.Desktop.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.Expanded = false;
            this.openFilePanel1.Desktop.Visible = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.openFilePanel1.Localization.AskBeforeDeleteDirectoryHeader = "Confirm";
            this.openFilePanel1.Localization.AskBeforeDeleteDirectoryMessage = "To confirm erasure press the YES key.";
            this.openFilePanel1.Localization.AskBeforeDeleteFileHeader = "Confirm";
            this.openFilePanel1.Localization.AskBeforeDeleteFileMessage = "To confirm erasure press the YES key.";
            this.openFilePanel1.Localization.Desktop = "Desktop";
            this.openFilePanel1.Localization.LabelCreateDirectory = "Create new Directory";
            this.openFilePanel1.Localization.LabelDeleteDirectory = "Delete Directory";
            this.openFilePanel1.Localization.LabelDeleteFile = "Delete File";
            this.openFilePanel1.Localization.LabelDetailsView = "Details";
            this.openFilePanel1.Localization.LabelFileFilter = "Filter";
            this.openFilePanel1.Localization.LabelFileName = "File(s)";
            this.openFilePanel1.Localization.LabelGoRedo = "Go Forward";
            this.openFilePanel1.Localization.LabelGoUndo = "Go Back";
            this.openFilePanel1.Localization.LabelGoUpward = "Go Upward";
            this.openFilePanel1.Localization.LabelLargeIconView = "Large Icons";
            this.openFilePanel1.Localization.LabelSmallIconView = "Small Icons";
            this.openFilePanel1.Localization.MyDocuments = "My Documents";
            this.openFilePanel1.Localization.MyMachine = "My Computer";
            this.openFilePanel1.Localization.NewDirectoryName = "New Directory";
            this.openFilePanel1.Localization.SpecialFolders = "Special Folders";
            this.openFilePanel1.Localization.TemplateFolders = "Custom Folders";
            this.openFilePanel1.Location = new System.Drawing.Point(4, 4);
            this.openFilePanel1.Margin = new System.Windows.Forms.Padding(4);
            this.openFilePanel1.Misc.AskBeforeDelete = true;
            this.openFilePanel1.Misc.CategoryPanelWidth = 230;
            this.openFilePanel1.Misc.FileFilter = "Assemblies(*.dll)|*.dll";
            this.openFilePanel1.Misc.FireSelectionChangedInsteadOfDoubleClick = true;
            this.openFilePanel1.Misc.SelectedCategory = NOTools.FileSystemDialogs.RootCategory.MyComputer;
            this.openFilePanel1.Misc.ShowCategoryPanel = true;
            this.openFilePanel1.Misc.ShowFilePanel = true;
            this.openFilePanel1.MyComputer.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.Expanded = false;
            this.openFilePanel1.MyComputer.ShowCDRomDrives = false;
            this.openFilePanel1.MyComputer.ShowFixedDrives = true;
            this.openFilePanel1.MyComputer.ShowNetworkDrives = false;
            this.openFilePanel1.MyComputer.ShowNoRootDirectoryDrives = false;
            this.openFilePanel1.MyComputer.ShowRamDrives = true;
            this.openFilePanel1.MyComputer.ShowRemovableDrives = false;
            this.openFilePanel1.MyComputer.ShowUnknownDrives = false;
            this.openFilePanel1.MyComputer.ShowUnreadyDrives = false;
            this.openFilePanel1.MyComputer.Visible = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.Expanded = false;
            this.openFilePanel1.MyDocuments.Visible = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.Name = "openFilePanel1";
            this.openFilePanel1.Size = new System.Drawing.Size(678, 386);
            this.openFilePanel1.SpecialFolders.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.SpecialFolders.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.SpecialFolders.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.SpecialFolders.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.SpecialFolders.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.SpecialFolders.Expanded = false;
            this.openFilePanel1.SpecialFolders.Visible = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.TabIndex = 0;
            this.openFilePanel1.TemplateFolders.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.TemplateFolders.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.TemplateFolders.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.TemplateFolders.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.TemplateFolders.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.TemplateFolders.Expanded = false;
            this.openFilePanel1.TemplateFolders.Visible = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.SelectionChanged += new NOTools.FileSystemDialogs.SelectionChangedEventHandler(this.openFilePanel1_SelectionChanged);
            // 
            // buttonOk
            // 
            this.buttonOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOk.Enabled = false;
            this.buttonOk.Location = new System.Drawing.Point(422, 435);
            this.buttonOk.Margin = new System.Windows.Forms.Padding(4);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(108, 28);
            this.buttonOk.TabIndex = 3;
            this.buttonOk.Text = "Ok";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Location = new System.Drawing.Point(550, 435);
            this.buttonCancel.Margin = new System.Windows.Forms.Padding(4);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(108, 28);
            this.buttonCancel.TabIndex = 4;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // ReferencesDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(692, 473);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ReferencesDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Add Reference";
            this.tabControl1.ResumeLayout(false);
            this.tabPageGac.ResumeLayout(false);
            this.tabPageGac.PerformLayout();
            this.tabPageFileSystem.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageGac;
        private System.Windows.Forms.TabPage tabPageFileSystem;
        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.Label labelNameFilter;
        private System.Windows.Forms.TextBox textBoxNameFilter;
        private NOTools.FileSystemDialogs.OpenFilePanel openFilePanel1;
    }
}