namespace NOToolsTests.CSharpTextEditor1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            NOTools.CSharpTextEditor.AssemblyReference assemblyReference1 = new NOTools.CSharpTextEditor.AssemblyReference();
            NOTools.CSharpTextEditor.AssemblyReference assemblyReference2 = new NOTools.CSharpTextEditor.AssemblyReference();
            NOTools.CSharpTextEditor.AssemblyReference assemblyReference3 = new NOTools.CSharpTextEditor.AssemblyReference();
            NOTools.CSharpTextEditor.AssemblyReference assemblyReference4 = new NOTools.CSharpTextEditor.AssemblyReference();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.StripButtonNew = new System.Windows.Forms.ToolStripButton();
            this.StripButtonOpen = new System.Windows.Forms.ToolStripButton();
            this.StripButtonSave = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.StripButtonRun = new System.Windows.Forms.ToolStripButton();
            this.StripButtonCompile = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.StripButtonAbout = new System.Windows.Forms.ToolStripButton();
            this.codeEditorControl1 = new NOTools.CSharpTextEditor.CodeEditorControl();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StripButtonNew,
            this.StripButtonOpen,
            this.StripButtonSave,
            this.toolStripSeparator,
            this.StripButtonRun,
            this.StripButtonCompile,
            this.toolStripSeparator1,
            this.StripButtonAbout});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(160, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // StripButtonNew
            // 
            this.StripButtonNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonNew.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonNew.Image")));
            this.StripButtonNew.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonNew.Name = "StripButtonNew";
            this.StripButtonNew.Size = new System.Drawing.Size(23, 22);
            this.StripButtonNew.Text = "New";
            this.StripButtonNew.Click += new System.EventHandler(this.StripButtonNew_Click);
            // 
            // StripButtonOpen
            // 
            this.StripButtonOpen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonOpen.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonOpen.Image")));
            this.StripButtonOpen.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonOpen.Name = "StripButtonOpen";
            this.StripButtonOpen.Size = new System.Drawing.Size(23, 22);
            this.StripButtonOpen.Text = "Open Code";
            this.StripButtonOpen.Click += new System.EventHandler(this.StripButtonOpen_Click);
            // 
            // StripButtonSave
            // 
            this.StripButtonSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonSave.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonSave.Image")));
            this.StripButtonSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonSave.Name = "StripButtonSave";
            this.StripButtonSave.Size = new System.Drawing.Size(23, 22);
            this.StripButtonSave.Text = "Save Code";
            this.StripButtonSave.Click += new System.EventHandler(this.StripButtonSave_Click);
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            this.toolStripSeparator.Size = new System.Drawing.Size(6, 25);
            // 
            // StripButtonRun
            // 
            this.StripButtonRun.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonRun.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonRun.Image")));
            this.StripButtonRun.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonRun.Name = "StripButtonRun";
            this.StripButtonRun.Size = new System.Drawing.Size(23, 22);
            this.StripButtonRun.Text = "Run Code";
            this.StripButtonRun.Click += new System.EventHandler(this.StripButtonRun_Click);
            // 
            // StripButtonCompile
            // 
            this.StripButtonCompile.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonCompile.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonCompile.Image")));
            this.StripButtonCompile.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonCompile.Name = "StripButtonCompile";
            this.StripButtonCompile.Size = new System.Drawing.Size(23, 22);
            this.StripButtonCompile.Text = "Compile Code";
            this.StripButtonCompile.Click += new System.EventHandler(this.StripButtonCompile_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // StripButtonAbout
            // 
            this.StripButtonAbout.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.StripButtonAbout.Image = ((System.Drawing.Image)(resources.GetObject("StripButtonAbout.Image")));
            this.StripButtonAbout.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripButtonAbout.Name = "StripButtonAbout";
            this.StripButtonAbout.Size = new System.Drawing.Size(23, 22);
            this.StripButtonAbout.Text = "About";
            this.StripButtonAbout.Click += new System.EventHandler(this.StripButtonAbout_Click);
            // 
            // codeEditorControl1
            // 
            this.codeEditorControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.codeEditorControl1.BackColor = System.Drawing.Color.Black;
            this.codeEditorControl1.CompileRequestOptions.CompileRequestKey = NOTools.CSharpTextEditor.Key.F6;
            this.codeEditorControl1.CompileRequestOptions.Enabled = true;
            this.codeEditorControl1.CompileRequestOptions.RunRequestKey = NOTools.CSharpTextEditor.Key.F5;
            this.codeEditorControl1.EnableFolding = true;
            this.codeEditorControl1.ErrorPanelSettings.AllowPanel = true;
            this.codeEditorControl1.ErrorPanelSettings.BackColor = System.Drawing.Color.LightSteelBlue;
            this.codeEditorControl1.ErrorPanelSettings.ErrorColumnHeader = "Error";
            this.codeEditorControl1.ErrorPanelSettings.ForeColor = System.Drawing.SystemColors.ControlText;
            this.codeEditorControl1.ErrorPanelSettings.Header = "Errors";
            this.codeEditorControl1.ErrorPanelSettings.LineColumnHeader = "Line";
            this.codeEditorControl1.ErrorPanelSettings.LineInfoFormatString = "Current Line:{0} Position:{1}";
            this.codeEditorControl1.ErrorPanelSettings.PanelOpen = false;
            this.codeEditorControl1.Location = new System.Drawing.Point(0, 28);
            this.codeEditorControl1.MinimumSize = new System.Drawing.Size(200, 200);
            this.codeEditorControl1.Name = "codeEditorControl1";
            this.codeEditorControl1.PersistencePath = "..\\..\\..\\Persistance Cache Debug";
            this.codeEditorControl1.ReferencePanelSettings.AddTitle = "Add";
            this.codeEditorControl1.ReferencePanelSettings.AllowAddRemoveReferences = true;
            this.codeEditorControl1.ReferencePanelSettings.AllowPanel = true;
            this.codeEditorControl1.ReferencePanelSettings.BackColor = System.Drawing.Color.LightSteelBlue;
            this.codeEditorControl1.ReferencePanelSettings.CancelButtonTitle = "Cancel";
            this.codeEditorControl1.ReferencePanelSettings.DialogTitle = "Choose Reference";
            this.codeEditorControl1.ReferencePanelSettings.FileSystemTitle = "FileSystem";
            this.codeEditorControl1.ReferencePanelSettings.ForeColor = System.Drawing.SystemColors.ControlText;
            this.codeEditorControl1.ReferencePanelSettings.GACTitle = "GAC";
            this.codeEditorControl1.ReferencePanelSettings.Header = "References";
            this.codeEditorControl1.ReferencePanelSettings.OkButtonTitle = "Ok";
            this.codeEditorControl1.ReferencePanelSettings.PanelOpen = true;
            this.codeEditorControl1.ReferencePanelSettings.RemoveTitle = "Remove";
            assemblyReference1.IsExe = false;
            assemblyReference1.Name = "System";
            assemblyReference1.Path = "";
            assemblyReference2.IsExe = false;
            assemblyReference2.Name = "System.Drawing";
            assemblyReference2.Path = "";
            assemblyReference3.IsExe = false;
            assemblyReference3.Name = "System.Windows.Forms";
            assemblyReference3.Path = "";
            assemblyReference4.IsExe = true;
            assemblyReference4.Name = "NOToolsTests.CSharpTextEditor1";
            assemblyReference4.Path = "";
            this.codeEditorControl1.References.Add(assemblyReference1);
            this.codeEditorControl1.References.Add(assemblyReference2);
            this.codeEditorControl1.References.Add(assemblyReference3);
            this.codeEditorControl1.References.Add(assemblyReference4);
            this.codeEditorControl1.ShowLineNumbers = true;
            this.codeEditorControl1.Size = new System.Drawing.Size(562, 367);
            this.codeEditorControl1.TabIndex = 1;
            this.codeEditorControl1.PersistanceResolve += new NOTools.CSharpTextEditor.PersistanceResolveEventHandler(this.codeEditorControl1_PersistanceResolve);
            this.codeEditorControl1.CompileRequest += new NOTools.CSharpTextEditor.CompileRequestHandler(this.codeEditorControl1_CompileRequest);
            this.codeEditorControl1.RunRequest += new NOTools.CSharpTextEditor.CompileRequestHandler(this.codeEditorControl1_RunRequest);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 393);
            this.Controls.Add(this.codeEditorControl1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NOToolsTests.CSharpTextEditor1";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton StripButtonNew;
        private System.Windows.Forms.ToolStripButton StripButtonOpen;
        private System.Windows.Forms.ToolStripButton StripButtonSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
        private System.Windows.Forms.ToolStripButton StripButtonRun;
        private System.Windows.Forms.ToolStripButton StripButtonCompile;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton StripButtonAbout;
        private NOTools.CSharpTextEditor.CodeEditorControl codeEditorControl1;
    }
}

