namespace NOTools.DeveloperAddin.UI
{
    partial class CommandPane
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
            this.gridCommands = new System.Windows.Forms.DataGridView();
            this.columnName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.columnReady = new System.Windows.Forms.DataGridViewImageColumn();
            this.columnAction = new System.Windows.Forms.DataGridViewLinkColumn();
            this.buttonAddNew = new System.Windows.Forms.Button();
            this.buttonExecuteCommand = new System.Windows.Forms.Button();
            this.panelCommands = new System.Windows.Forms.Panel();
            this.panelCodeEditor = new System.Windows.Forms.Panel();
            this.codeEditor = new NOTools.CSharpTextEditor.CodeEditorControl();
            this.buttonSkip = new System.Windows.Forms.Button();
            this.buttonCompileAndExecute = new System.Windows.Forms.Button();
            this.buttonApply = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.gridCommands)).BeginInit();
            this.panelCommands.SuspendLayout();
            this.panelCodeEditor.SuspendLayout();
            this.SuspendLayout();
            // 
            // gridCommands
            // 
            this.gridCommands.AllowUserToResizeRows = false;
            this.gridCommands.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gridCommands.BackgroundColor = System.Drawing.Color.LightSteelBlue;
            this.gridCommands.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridCommands.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.columnName,
            this.columnReady,
            this.columnAction});
            this.gridCommands.GridColor = System.Drawing.Color.Black;
            this.gridCommands.Location = new System.Drawing.Point(4, 4);
            this.gridCommands.Margin = new System.Windows.Forms.Padding(4);
            this.gridCommands.Name = "gridCommands";
            this.gridCommands.RowHeadersVisible = false;
            this.gridCommands.Size = new System.Drawing.Size(402, 187);
            this.gridCommands.TabIndex = 10;
            // 
            // columnName
            // 
            this.columnName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.columnName.DataPropertyName = "Name";
            this.columnName.HeaderText = "Command";
            this.columnName.Name = "columnName";
            // 
            // columnReady
            // 
            this.columnReady.DataPropertyName = "Ready";
            this.columnReady.HeaderText = "Ready";
            this.columnReady.Name = "columnReady";
            this.columnReady.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.columnReady.Width = 50;
            // 
            // columnAction
            // 
            this.columnAction.HeaderText = "Edit / Delete";
            this.columnAction.Name = "columnAction";
            this.columnAction.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // buttonAddNew
            // 
            this.buttonAddNew.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonAddNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonAddNew.Location = new System.Drawing.Point(4, 190);
            this.buttonAddNew.Margin = new System.Windows.Forms.Padding(4);
            this.buttonAddNew.Name = "buttonAddNew";
            this.buttonAddNew.Size = new System.Drawing.Size(118, 30);
            this.buttonAddNew.TabIndex = 9;
            this.buttonAddNew.Text = "Add New";
            this.buttonAddNew.UseVisualStyleBackColor = true;
            this.buttonAddNew.Click += new System.EventHandler(this.buttonAddNew_Click);
            // 
            // buttonExecuteCommand
            // 
            this.buttonExecuteCommand.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonExecuteCommand.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonExecuteCommand.Location = new System.Drawing.Point(121, 190);
            this.buttonExecuteCommand.Margin = new System.Windows.Forms.Padding(4);
            this.buttonExecuteCommand.Name = "buttonExecuteCommand";
            this.buttonExecuteCommand.Size = new System.Drawing.Size(285, 30);
            this.buttonExecuteCommand.TabIndex = 8;
            this.buttonExecuteCommand.Text = "Execute Command";
            this.buttonExecuteCommand.UseVisualStyleBackColor = true;
            this.buttonExecuteCommand.Click += new System.EventHandler(this.buttonExecuteCommand_Click);
            // 
            // panelCommands
            // 
            this.panelCommands.Controls.Add(this.gridCommands);
            this.panelCommands.Controls.Add(this.buttonAddNew);
            this.panelCommands.Controls.Add(this.buttonExecuteCommand);
            this.panelCommands.Location = new System.Drawing.Point(3, 3);
            this.panelCommands.Name = "panelCommands";
            this.panelCommands.Size = new System.Drawing.Size(410, 224);
            this.panelCommands.TabIndex = 11;
            // 
            // panelCodeEditor
            // 
            this.panelCodeEditor.Controls.Add(this.buttonApply);
            this.panelCodeEditor.Controls.Add(this.buttonCompileAndExecute);
            this.panelCodeEditor.Controls.Add(this.codeEditor);
            this.panelCodeEditor.Controls.Add(this.buttonSkip);
            this.panelCodeEditor.Location = new System.Drawing.Point(3, 231);
            this.panelCodeEditor.Name = "panelCodeEditor";
            this.panelCodeEditor.Size = new System.Drawing.Size(410, 244);
            this.panelCodeEditor.TabIndex = 12;
            // 
            // codeEditor
            // 
            this.codeEditor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.codeEditor.Location = new System.Drawing.Point(0, 0);
            this.codeEditor.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.codeEditor.Name = "codeEditor";
            this.codeEditor.Size = new System.Drawing.Size(408, 210);
            this.codeEditor.TabIndex = 0;
            // 
            // buttonSkip
            // 
            this.buttonSkip.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonSkip.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSkip.Location = new System.Drawing.Point(0, 210);
            this.buttonSkip.Margin = new System.Windows.Forms.Padding(4);
            this.buttonSkip.Name = "buttonSkip";
            this.buttonSkip.Size = new System.Drawing.Size(118, 30);
            this.buttonSkip.TabIndex = 11;
            this.buttonSkip.Text = "Skip";
            this.buttonSkip.UseVisualStyleBackColor = true;
            // 
            // buttonCompileAndExecute
            // 
            this.buttonCompileAndExecute.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCompileAndExecute.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonCompileAndExecute.Location = new System.Drawing.Point(232, 210);
            this.buttonCompileAndExecute.Margin = new System.Windows.Forms.Padding(4);
            this.buttonCompileAndExecute.Name = "buttonCompileAndExecute";
            this.buttonCompileAndExecute.Size = new System.Drawing.Size(177, 30);
            this.buttonCompileAndExecute.TabIndex = 10;
            this.buttonCompileAndExecute.Text = "Compile && Execute";
            this.buttonCompileAndExecute.UseVisualStyleBackColor = true;
            // 
            // buttonApply
            // 
            this.buttonApply.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonApply.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonApply.Location = new System.Drawing.Point(118, 210);
            this.buttonApply.Margin = new System.Windows.Forms.Padding(4);
            this.buttonApply.Name = "buttonApply";
            this.buttonApply.Size = new System.Drawing.Size(118, 30);
            this.buttonApply.TabIndex = 12;
            this.buttonApply.Text = "Apply && Close";
            this.buttonApply.UseVisualStyleBackColor = true;
            // 
            // CommandPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.panelCodeEditor);
            this.Controls.Add(this.panelCommands);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "CommandPane";
            this.Size = new System.Drawing.Size(416, 478);
            ((System.ComponentModel.ISupportInitialize)(this.gridCommands)).EndInit();
            this.panelCommands.ResumeLayout(false);
            this.panelCodeEditor.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView gridCommands;
        private System.Windows.Forms.Button buttonAddNew;
        private System.Windows.Forms.Button buttonExecuteCommand;
        private System.Windows.Forms.DataGridViewTextBoxColumn columnName;
        private System.Windows.Forms.DataGridViewImageColumn columnReady;
        private System.Windows.Forms.DataGridViewLinkColumn columnAction;
        private System.Windows.Forms.Panel panelCommands;
        private System.Windows.Forms.Panel panelCodeEditor;
        private CSharpTextEditor.CodeEditorControl codeEditor;
        private System.Windows.Forms.Button buttonSkip;
        private System.Windows.Forms.Button buttonCompileAndExecute;
        private System.Windows.Forms.Button buttonApply;
    }
}
