namespace NOTools.FileSystemDialogs
{
    partial class OpenFilePanelDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OpenFilePanelDialog));
            this.ButtonCancel = new System.Windows.Forms.Button();
            this.ButtonSelect = new System.Windows.Forms.Button();
            this.InnerOpenFilePanel = new NOTools.FileSystemDialogs.OpenFilePanel();
            this.SuspendLayout();
            // 
            // ButtonCancel
            // 
            this.ButtonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.ButtonCancel.Image = ((System.Drawing.Image)(resources.GetObject("ButtonCancel.Image")));
            this.ButtonCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ButtonCancel.Location = new System.Drawing.Point(503, 298);
            this.ButtonCancel.Name = "ButtonCancel";
            this.ButtonCancel.Size = new System.Drawing.Size(102, 28);
            this.ButtonCancel.TabIndex = 1;
            this.ButtonCancel.Text = "Cancel";
            this.ButtonCancel.UseVisualStyleBackColor = true;
            this.ButtonCancel.Click += new System.EventHandler(this.ButtonCancel_Click);
            // 
            // ButtonSelect
            // 
            this.ButtonSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ButtonSelect.Enabled = false;
            this.ButtonSelect.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSelect.Image")));
            this.ButtonSelect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ButtonSelect.Location = new System.Drawing.Point(383, 298);
            this.ButtonSelect.Name = "ButtonSelect";
            this.ButtonSelect.Size = new System.Drawing.Size(103, 28);
            this.ButtonSelect.TabIndex = 2;
            this.ButtonSelect.Text = "Open";
            this.ButtonSelect.UseVisualStyleBackColor = true;
            this.ButtonSelect.Click += new System.EventHandler(this.ButtonSelect_Click);
            // 
            // InnerOpenFilePanel
            // 
            this.InnerOpenFilePanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.InnerOpenFilePanel.Default.AllowAddFolders = false;
            this.InnerOpenFilePanel.Default.AllowBrowseFolders = true;
            this.InnerOpenFilePanel.Default.AllowDeleteFiles = false;
            this.InnerOpenFilePanel.Default.AllowDeleteFolders = false;
            this.InnerOpenFilePanel.Default.AllowMultipleSelect = false;
            this.InnerOpenFilePanel.Default.Visible = true;
            this.InnerOpenFilePanel.Desktop.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.Desktop.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.Desktop.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.Desktop.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.Desktop.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.Desktop.Expanded = false;
            this.InnerOpenFilePanel.Desktop.Visible = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.Localization.AskBeforeDeleteDirectoryHeader = "Confirm";
            this.InnerOpenFilePanel.Localization.AskBeforeDeleteDirectoryMessage = "To confirm erasure press the YES key.";
            this.InnerOpenFilePanel.Localization.AskBeforeDeleteFileHeader = "Confirm";
            this.InnerOpenFilePanel.Localization.AskBeforeDeleteFileMessage = "To confirm erasure press the YES key.";
            this.InnerOpenFilePanel.Localization.Desktop = "Desktop";
            this.InnerOpenFilePanel.Localization.LabelCreateDirectory = "Create new Directory";
            this.InnerOpenFilePanel.Localization.LabelDeleteDirectory = "Delete Directory";
            this.InnerOpenFilePanel.Localization.LabelDeleteFile = "Delete File";
            this.InnerOpenFilePanel.Localization.LabelDetailsView = "Details";
            this.InnerOpenFilePanel.Localization.LabelFileFilter = "Filter";
            this.InnerOpenFilePanel.Localization.LabelFileName = "File(s)";
            this.InnerOpenFilePanel.Localization.LabelGoRedo = "Go Forward";
            this.InnerOpenFilePanel.Localization.LabelGoUndo = "Go Back";
            this.InnerOpenFilePanel.Localization.LabelGoUpward = "Go Upward";
            this.InnerOpenFilePanel.Localization.LabelLargeIconView = "Large Icons";
            this.InnerOpenFilePanel.Localization.LabelSmallIconView = "Small Icons";
            this.InnerOpenFilePanel.Localization.MyDocuments = "My Documents";
            this.InnerOpenFilePanel.Localization.MyMachine = "My Computer";
            this.InnerOpenFilePanel.Localization.NewDirectoryName = "New Directory";
            this.InnerOpenFilePanel.Localization.SpecialFolders = "Special Folders";
            this.InnerOpenFilePanel.Localization.TemplateFolders = "Custom Folders";
            this.InnerOpenFilePanel.Location = new System.Drawing.Point(0, 0);
            this.InnerOpenFilePanel.Misc.AskBeforeDelete = true;
            this.InnerOpenFilePanel.Misc.CategoryPanelWidth = 230;
            this.InnerOpenFilePanel.Misc.FileFilter = "";
            this.InnerOpenFilePanel.Misc.FireSelectionChangedInsteadOfDoubleClick = false;
            this.InnerOpenFilePanel.Misc.SelectedCategory = NOTools.FileSystemDialogs.RootCategory.Desktop;
            this.InnerOpenFilePanel.Misc.ShowCategoryPanel = true;
            this.InnerOpenFilePanel.Misc.ShowFilePanel = true;
            this.InnerOpenFilePanel.MyComputer.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyComputer.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyComputer.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyComputer.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyComputer.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyComputer.Expanded = false;
            this.InnerOpenFilePanel.MyComputer.ShowCDRomDrives = true;
            this.InnerOpenFilePanel.MyComputer.ShowFixedDrives = true;
            this.InnerOpenFilePanel.MyComputer.ShowNetworkDrives = true;
            this.InnerOpenFilePanel.MyComputer.ShowNoRootDirectoryDrives = true;
            this.InnerOpenFilePanel.MyComputer.ShowRamDrives = true;
            this.InnerOpenFilePanel.MyComputer.ShowRemovableDrives = true;
            this.InnerOpenFilePanel.MyComputer.ShowUnknownDrives = true;
            this.InnerOpenFilePanel.MyComputer.ShowUnreadyDrives = true;
            this.InnerOpenFilePanel.MyComputer.Visible = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyDocuments.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyDocuments.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyDocuments.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyDocuments.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyDocuments.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.MyDocuments.Expanded = false;
            this.InnerOpenFilePanel.MyDocuments.Visible = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.Name = "InnerOpenFilePanel";
            this.InnerOpenFilePanel.Size = new System.Drawing.Size(636, 292);
            this.InnerOpenFilePanel.SpecialFolders.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.InnerOpenFilePanel.SpecialFolders.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.SpecialFolders.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.SpecialFolders.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.InnerOpenFilePanel.SpecialFolders.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.SpecialFolders.Expanded = false;
            this.InnerOpenFilePanel.SpecialFolders.Visible = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.InnerOpenFilePanel.TabIndex = 0;
            this.InnerOpenFilePanel.TemplateFolders.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.InnerOpenFilePanel.TemplateFolders.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.TemplateFolders.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.TemplateFolders.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.InnerOpenFilePanel.TemplateFolders.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.InnerOpenFilePanel.TemplateFolders.Expanded = false;
            this.InnerOpenFilePanel.TemplateFolders.Visible = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.InnerOpenFilePanel.FileDoubleClick += new NOTools.FileSystemDialogs.FileDoubleClickEventHandler(this.InnerOpenFilePanel_FileDoubleClick);
            this.InnerOpenFilePanel.SelectionChanged += new NOTools.FileSystemDialogs.SelectionChangedEventHandler(this.InnerOpenFilePanel_SelectionChanged);
            // 
            // OpenFilePanelDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.ButtonCancel;
            this.ClientSize = new System.Drawing.Size(636, 334);
            this.Controls.Add(this.ButtonSelect);
            this.Controls.Add(this.ButtonCancel);
            this.Controls.Add(this.InnerOpenFilePanel);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Name = "OpenFilePanelDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Open file";
            this.ResumeLayout(false);

        }

        #endregion

        private OpenFilePanel InnerOpenFilePanel;
        private System.Windows.Forms.Button ButtonCancel;
        private System.Windows.Forms.Button ButtonSelect;
    }
}