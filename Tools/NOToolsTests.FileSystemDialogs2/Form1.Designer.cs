namespace NOToolsTests.FileSystemDialogs2
{
    partial class Form1
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

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            NOTools.FileSystemDialogs.TemplateFolderDescription templateFolderDescription1 = new NOTools.FileSystemDialogs.TemplateFolderDescription();
            NOTools.FileSystemDialogs.TemplateFolderDescription templateFolderDescription2 = new NOTools.FileSystemDialogs.TemplateFolderDescription();
            this.ButtonSelectFile = new System.Windows.Forms.Button();
            this.openFilePanel1 = new NOTools.FileSystemDialogs.OpenFilePanel();
            this.SuspendLayout();
            // 
            // ButtonSelectFile
            // 
            this.ButtonSelectFile.Enabled = false;
            this.ButtonSelectFile.Location = new System.Drawing.Point(493, 341);
            this.ButtonSelectFile.Name = "ButtonSelectFile";
            this.ButtonSelectFile.Size = new System.Drawing.Size(87, 25);
            this.ButtonSelectFile.TabIndex = 1;
            this.ButtonSelectFile.Text = "Select File";
            this.ButtonSelectFile.UseVisualStyleBackColor = true;
            this.ButtonSelectFile.Click += new System.EventHandler(this.ButtonSelectFile_Click);
            // 
            // openFilePanel1
            // 
            this.openFilePanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.openFilePanel1.Default.AllowAddFolders = false;
            this.openFilePanel1.Default.AllowBrowseFolders = true;
            this.openFilePanel1.Default.AllowDeleteFiles = false;
            this.openFilePanel1.Default.AllowDeleteFolders = false;
            this.openFilePanel1.Default.AllowMultipleSelect = false;
            this.openFilePanel1.Default.Visible = true;
            this.openFilePanel1.Desktop.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Desktop.Expanded = false;
            this.openFilePanel1.Desktop.Visible = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Localization.AskBeforeDeleteDirectoryHeader = "Bestätigen";
            this.openFilePanel1.Localization.AskBeforeDeleteDirectoryMessage = "Ausgewählte Ordner löschen?";
            this.openFilePanel1.Localization.AskBeforeDeleteFileHeader = "Bestätigen";
            this.openFilePanel1.Localization.AskBeforeDeleteFileMessage = "Ausgewählte Dateien löschen?";
            this.openFilePanel1.Localization.Desktop = "Desktop";
            this.openFilePanel1.Localization.LabelCreateDirectory = "Neuen Ordner erstellen";
            this.openFilePanel1.Localization.LabelDeleteDirectory = "Ordner löschen";
            this.openFilePanel1.Localization.LabelDeleteFile = "Datei löschen";
            this.openFilePanel1.Localization.LabelDetailsView = "Detailansicht";
            this.openFilePanel1.Localization.LabelFileFilter = "Filter";
            this.openFilePanel1.Localization.LabelFileName = "Datei";
            this.openFilePanel1.Localization.LabelGoRedo = "Vor";
            this.openFilePanel1.Localization.LabelGoUndo = "Zurück";
            this.openFilePanel1.Localization.LabelGoUpward = "Nach oben";
            this.openFilePanel1.Localization.LabelLargeIconView = "Grosse Symbole";
            this.openFilePanel1.Localization.LabelSmallIconView = "Kleine Symbole";
            this.openFilePanel1.Localization.MyDocuments = "Eigene Dateien";
            this.openFilePanel1.Localization.MyMachine = "Computer";
            this.openFilePanel1.Localization.NewDirectoryName = "Neuer Ordner";
            this.openFilePanel1.Localization.SpecialFolders = "Registrierte Ordner";
            this.openFilePanel1.Localization.TemplateFolders = "Auswahl";
            this.openFilePanel1.Location = new System.Drawing.Point(0, 0);
            this.openFilePanel1.Misc.AskBeforeDelete = true;
            this.openFilePanel1.Misc.CategoryPanelWidth = 230;
            this.openFilePanel1.Misc.FileFilter = "";
            this.openFilePanel1.Misc.FireSelectionChangedInsteadOfDoubleClick = false;
            this.openFilePanel1.Misc.SelectedCategory = NOTools.FileSystemDialogs.RootCategory.Desktop;
            this.openFilePanel1.Misc.ShowCategoryPanel = true;
            this.openFilePanel1.Misc.ShowFilePanel = true;
            this.openFilePanel1.MyComputer.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyComputer.Expanded = false;
            this.openFilePanel1.MyComputer.ShowCDRomDrives = true;
            this.openFilePanel1.MyComputer.ShowFixedDrives = true;
            this.openFilePanel1.MyComputer.ShowNetworkDrives = true;
            this.openFilePanel1.MyComputer.ShowNoRootDirectoryDrives = true;
            this.openFilePanel1.MyComputer.ShowRamDrives = true;
            this.openFilePanel1.MyComputer.ShowRemovableDrives = true;
            this.openFilePanel1.MyComputer.ShowUnknownDrives = true;
            this.openFilePanel1.MyComputer.ShowUnreadyDrives = true;
            this.openFilePanel1.MyComputer.Visible = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.MyDocuments.Expanded = false;
            this.openFilePanel1.MyDocuments.Visible = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.Name = "openFilePanel1";
            this.openFilePanel1.Size = new System.Drawing.Size(609, 328);
            this.openFilePanel1.SpecialFolders.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.SpecialFolders.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.SpecialFolders.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.SpecialFolders.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.SpecialFolders.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.SpecialFolders.Expanded = false;
            this.openFilePanel1.SpecialFolders.Visible = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.TabIndex = 0;
            this.openFilePanel1.TemplateFolders.AllowAddFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.TemplateFolders.AllowBrowseFolders = NOTools.FileSystemDialogs.DefaultBoolean.True;
            this.openFilePanel1.TemplateFolders.AllowDeleteFiles = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.TemplateFolders.AllowDeleteFolders = NOTools.FileSystemDialogs.DefaultBoolean.False;
            this.openFilePanel1.TemplateFolders.AllowMultipleSelect = NOTools.FileSystemDialogs.DefaultBoolean.Default;
            this.openFilePanel1.TemplateFolders.Expanded = false;
            templateFolderDescription1.DisplayName = "Windows";
            templateFolderDescription1.Path = "C:\\Windows";
            templateFolderDescription2.DisplayName = "Root";
            templateFolderDescription2.Path = "C:\\";
            this.openFilePanel1.TemplateFolders.FolderTemplates.Add(templateFolderDescription1);
            this.openFilePanel1.TemplateFolders.FolderTemplates.Add(templateFolderDescription2);
            this.openFilePanel1.TemplateFolders.Visible = NOTools.FileSystemDialogs.DefaultBoolean.True;
            this.openFilePanel1.FileDoubleClick += new NOTools.FileSystemDialogs.FileDoubleClickEventHandler(this.openFilePanel1_FileDoubleClick);
            this.openFilePanel1.SelectionChanged += new NOTools.FileSystemDialogs.SelectionChangedEventHandler(this.openFilePanel1_SelectionChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(609, 378);
            this.Controls.Add(this.ButtonSelectFile);
            this.Controls.Add(this.openFilePanel1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NOToolsTests.FileSystemDialogs2";
            this.ResumeLayout(false);

        }

        #endregion

        private NOTools.FileSystemDialogs.OpenFilePanel openFilePanel1;
        private System.Windows.Forms.Button ButtonSelectFile;
    }
}

