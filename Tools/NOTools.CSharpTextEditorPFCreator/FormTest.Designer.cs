namespace NOTools.CSharpTextEditorPFCreator
{
    partial class FormTest
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
            NOTools.CSharpTextEditor.AssemblyReference assemblyReference1 = new NOTools.CSharpTextEditor.AssemblyReference();
            NOTools.CSharpTextEditor.AssemblyReference assemblyReference2 = new NOTools.CSharpTextEditor.AssemblyReference();
            NOTools.CSharpTextEditor.AssemblyReference assemblyReference3 = new NOTools.CSharpTextEditor.AssemblyReference();
            this.groupBoxSettings = new System.Windows.Forms.GroupBox();
            this.checkBoxAsync = new System.Windows.Forms.CheckBox();
            this.buttonDelete = new System.Windows.Forms.Button();
            this.buttonAdd = new System.Windows.Forms.Button();
            this.buttonStart = new System.Windows.Forms.Button();
            this.listViewAssemblies = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxResultFolder = new System.Windows.Forms.TextBox();
            this.buttonChooseFolder = new System.Windows.Forms.Button();
            this.codeEditorControl1 = new NOTools.CSharpTextEditor.CodeEditorControl();
            this.groupBoxSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxSettings
            // 
            this.groupBoxSettings.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxSettings.Controls.Add(this.checkBoxAsync);
            this.groupBoxSettings.Controls.Add(this.buttonDelete);
            this.groupBoxSettings.Controls.Add(this.buttonAdd);
            this.groupBoxSettings.Controls.Add(this.buttonStart);
            this.groupBoxSettings.Controls.Add(this.listViewAssemblies);
            this.groupBoxSettings.Controls.Add(this.label1);
            this.groupBoxSettings.Controls.Add(this.textBoxResultFolder);
            this.groupBoxSettings.Controls.Add(this.buttonChooseFolder);
            this.groupBoxSettings.Location = new System.Drawing.Point(0, 403);
            this.groupBoxSettings.Name = "groupBoxSettings";
            this.groupBoxSettings.Size = new System.Drawing.Size(703, 181);
            this.groupBoxSettings.TabIndex = 6;
            this.groupBoxSettings.TabStop = false;
            // 
            // checkBoxAsync
            // 
            this.checkBoxAsync.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBoxAsync.AutoSize = true;
            this.checkBoxAsync.Location = new System.Drawing.Point(254, 150);
            this.checkBoxAsync.Name = "checkBoxAsync";
            this.checkBoxAsync.Size = new System.Drawing.Size(55, 17);
            this.checkBoxAsync.TabIndex = 13;
            this.checkBoxAsync.Text = "Async";
            this.checkBoxAsync.UseVisualStyleBackColor = true;
            // 
            // buttonDelete
            // 
            this.buttonDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonDelete.Enabled = false;
            this.buttonDelete.Location = new System.Drawing.Point(113, 145);
            this.buttonDelete.Name = "buttonDelete";
            this.buttonDelete.Size = new System.Drawing.Size(94, 24);
            this.buttonDelete.TabIndex = 12;
            this.buttonDelete.Text = "Delete";
            this.buttonDelete.UseVisualStyleBackColor = true;
            this.buttonDelete.Click += new System.EventHandler(this.buttonDelete_Click);
            // 
            // buttonAdd
            // 
            this.buttonAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonAdd.Location = new System.Drawing.Point(13, 145);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(94, 24);
            this.buttonAdd.TabIndex = 11;
            this.buttonAdd.Text = "Add";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
            // 
            // buttonStart
            // 
            this.buttonStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonStart.Location = new System.Drawing.Point(543, 145);
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.Size = new System.Drawing.Size(94, 24);
            this.buttonStart.TabIndex = 10;
            this.buttonStart.Text = "Set References";
            this.buttonStart.UseVisualStyleBackColor = true;
            this.buttonStart.Click += new System.EventHandler(this.buttonStart_Click);
            // 
            // listViewAssemblies
            // 
            this.listViewAssemblies.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewAssemblies.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this.listViewAssemblies.Location = new System.Drawing.Point(13, 45);
            this.listViewAssemblies.Name = "listViewAssemblies";
            this.listViewAssemblies.Size = new System.Drawing.Size(624, 94);
            this.listViewAssemblies.TabIndex = 9;
            this.listViewAssemblies.UseCompatibleStateImageBehavior = false;
            this.listViewAssemblies.View = System.Windows.Forms.View.Details;
            this.listViewAssemblies.SelectedIndexChanged += new System.EventHandler(this.listViewAssemblies_SelectedIndexChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Name";
            this.columnHeader1.Width = 154;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Path";
            this.columnHeader2.Width = 227;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Persistence Path";
            // 
            // textBoxResultFolder
            // 
            this.textBoxResultFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxResultFolder.Location = new System.Drawing.Point(103, 19);
            this.textBoxResultFolder.Name = "textBoxResultFolder";
            this.textBoxResultFolder.Size = new System.Drawing.Size(534, 20);
            this.textBoxResultFolder.TabIndex = 1;
            // 
            // buttonChooseFolder
            // 
            this.buttonChooseFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonChooseFolder.Location = new System.Drawing.Point(658, 16);
            this.buttonChooseFolder.Name = "buttonChooseFolder";
            this.buttonChooseFolder.Size = new System.Drawing.Size(35, 24);
            this.buttonChooseFolder.TabIndex = 0;
            this.buttonChooseFolder.Text = "...";
            this.buttonChooseFolder.UseVisualStyleBackColor = true;
            this.buttonChooseFolder.Click += new System.EventHandler(this.buttonChooseFolder_Click);
            // 
            // codeEditorControl1
            // 
            this.codeEditorControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.codeEditorControl1.BackColor = System.Drawing.Color.White;
            this.codeEditorControl1.CompileRequestOptions.CompileRequestKey = NOTools.CSharpTextEditor.Key.F5;
            this.codeEditorControl1.CompileRequestOptions.Enabled = true;
            this.codeEditorControl1.ErrorPanelSettings.AllowPanel = true;
            this.codeEditorControl1.ErrorPanelSettings.BackColor = System.Drawing.Color.LightSteelBlue;
            this.codeEditorControl1.ErrorPanelSettings.ErrorColumnHeader = "Error";
            this.codeEditorControl1.ErrorPanelSettings.ForeColor = System.Drawing.SystemColors.ControlText;
            this.codeEditorControl1.ErrorPanelSettings.Header = "Errors";
            this.codeEditorControl1.ErrorPanelSettings.LineColumnHeader = "Line";
            this.codeEditorControl1.ErrorPanelSettings.LineInfoFormatString = "Current Line:{0} Position:{1}";
            this.codeEditorControl1.ErrorPanelSettings.PanelOpen = false;
            this.codeEditorControl1.Location = new System.Drawing.Point(0, 0);
            this.codeEditorControl1.MinimumSize = new System.Drawing.Size(400, 300);
            this.codeEditorControl1.Name = "codeEditorControl1";
            this.codeEditorControl1.PersistencePath = "C:\\PF-Test\\PF-Files";
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
            assemblyReference1.Name = "System";
            assemblyReference1.Path = null;
            assemblyReference2.Name = "System.Drawing";
            assemblyReference2.Path = null;
            assemblyReference3.Name = "System.Windows.Forms";
            assemblyReference3.Path = null;
            this.codeEditorControl1.References.Add(assemblyReference1);
            this.codeEditorControl1.References.Add(assemblyReference2);
            this.codeEditorControl1.References.Add(assemblyReference3);
            this.codeEditorControl1.ShowLineNumbers = true;
            this.codeEditorControl1.Size = new System.Drawing.Size(703, 397);
            this.codeEditorControl1.TabIndex = 0;
            this.codeEditorControl1.CompileRequest += new NOTools.CSharpTextEditor.CompileRequestHandler(this.codeEditorControl1_CompileRequest);
            // 
            // FormTest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(705, 583);
            this.Controls.Add(this.groupBoxSettings);
            this.Controls.Add(this.codeEditorControl1);
            this.Name = "FormTest";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CSharpTextEditor PersistenceFile Creator Test";
            this.groupBoxSettings.ResumeLayout(false);
            this.groupBoxSettings.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private CSharpTextEditor.CodeEditorControl codeEditorControl1;
        private System.Windows.Forms.GroupBox groupBoxSettings;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxResultFolder;
        private System.Windows.Forms.Button buttonChooseFolder;
        private System.Windows.Forms.Button buttonDelete;
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.ListView listViewAssemblies;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.CheckBox checkBoxAsync;
    }
}