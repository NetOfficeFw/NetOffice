namespace NOToolsTests.CSharpTextEditor2
{
    partial class FormHost
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dataDisplayControl1 = new NOToolsTests.CSharpTextEditor2.DataDisplayControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dataDisplayControl2 = new NOToolsTests.CSharpTextEditor2.DataDisplayControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dataDisplayControl3 = new NOToolsTests.CSharpTextEditor2.DataDisplayControl();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.dataDisplayControl4 = new NOToolsTests.CSharpTextEditor2.DataDisplayControl();
            this.buttonResetData = new System.Windows.Forms.Button();
            this.buttonCancelChanges = new System.Windows.Forms.Button();
            this.buttonSaveChanges = new System.Windows.Forms.Button();
            this.buttonSimulateDatabaseAction = new System.Windows.Forms.Button();
            this.buttonDoUndo = new System.Windows.Forms.Button();
            this.buttonDoRedo = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(0, 48);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1053, 550);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.Selecting += new System.Windows.Forms.TabControlCancelEventHandler(this.tabControl1_Selecting);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataDisplayControl1);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1045, 521);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Persons1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataDisplayControl1
            // 
            this.dataDisplayControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataDisplayControl1.Location = new System.Drawing.Point(3, 3);
            this.dataDisplayControl1.Name = "dataDisplayControl1";
            this.dataDisplayControl1.Size = new System.Drawing.Size(1039, 515);
            this.dataDisplayControl1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dataDisplayControl2);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1045, 521);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Products1";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dataDisplayControl2
            // 
            this.dataDisplayControl2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataDisplayControl2.Location = new System.Drawing.Point(3, 3);
            this.dataDisplayControl2.Name = "dataDisplayControl2";
            this.dataDisplayControl2.Size = new System.Drawing.Size(759, 357);
            this.dataDisplayControl2.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dataDisplayControl3);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(1045, 521);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Persons2";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dataDisplayControl3
            // 
            this.dataDisplayControl3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataDisplayControl3.Location = new System.Drawing.Point(3, 3);
            this.dataDisplayControl3.Name = "dataDisplayControl3";
            this.dataDisplayControl3.Size = new System.Drawing.Size(759, 357);
            this.dataDisplayControl3.TabIndex = 0;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.dataDisplayControl4);
            this.tabPage4.Location = new System.Drawing.Point(4, 25);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(1045, 521);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Products2";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // dataDisplayControl4
            // 
            this.dataDisplayControl4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataDisplayControl4.Location = new System.Drawing.Point(3, 3);
            this.dataDisplayControl4.Name = "dataDisplayControl4";
            this.dataDisplayControl4.Size = new System.Drawing.Size(759, 357);
            this.dataDisplayControl4.TabIndex = 0;
            // 
            // buttonResetData
            // 
            this.buttonResetData.Location = new System.Drawing.Point(386, 1);
            this.buttonResetData.Name = "buttonResetData";
            this.buttonResetData.Size = new System.Drawing.Size(186, 30);
            this.buttonResetData.TabIndex = 9;
            this.buttonResetData.Text = "Reload Data";
            this.buttonResetData.UseVisualStyleBackColor = true;
            this.buttonResetData.Click += new System.EventHandler(this.buttonResetData_Click);
            // 
            // buttonCancelChanges
            // 
            this.buttonCancelChanges.Enabled = false;
            this.buttonCancelChanges.Location = new System.Drawing.Point(194, 1);
            this.buttonCancelChanges.Name = "buttonCancelChanges";
            this.buttonCancelChanges.Size = new System.Drawing.Size(186, 30);
            this.buttonCancelChanges.TabIndex = 8;
            this.buttonCancelChanges.Text = "Cancel Local Changes";
            this.buttonCancelChanges.UseVisualStyleBackColor = true;
            this.buttonCancelChanges.Click += new System.EventHandler(this.buttonCancelChanges_Click);
            // 
            // buttonSaveChanges
            // 
            this.buttonSaveChanges.Enabled = false;
            this.buttonSaveChanges.Location = new System.Drawing.Point(2, 1);
            this.buttonSaveChanges.Name = "buttonSaveChanges";
            this.buttonSaveChanges.Size = new System.Drawing.Size(186, 30);
            this.buttonSaveChanges.TabIndex = 7;
            this.buttonSaveChanges.Text = "Save Local Changes";
            this.buttonSaveChanges.UseVisualStyleBackColor = true;
            this.buttonSaveChanges.Click += new System.EventHandler(this.buttonSaveChanges_Click);
            // 
            // buttonSimulateDatabaseAction
            // 
            this.buttonSimulateDatabaseAction.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSimulateDatabaseAction.Location = new System.Drawing.Point(863, 1);
            this.buttonSimulateDatabaseAction.Name = "buttonSimulateDatabaseAction";
            this.buttonSimulateDatabaseAction.Size = new System.Drawing.Size(186, 30);
            this.buttonSimulateDatabaseAction.TabIndex = 10;
            this.buttonSimulateDatabaseAction.Text = "Simulate DB Action";
            this.buttonSimulateDatabaseAction.UseVisualStyleBackColor = true;
            this.buttonSimulateDatabaseAction.Click += new System.EventHandler(this.buttonSimulateDatabaseAction_Click);
            // 
            // buttonDoUndo
            // 
            this.buttonDoUndo.Enabled = false;
            this.buttonDoUndo.Location = new System.Drawing.Point(594, 1);
            this.buttonDoUndo.Name = "buttonDoUndo";
            this.buttonDoUndo.Size = new System.Drawing.Size(83, 30);
            this.buttonDoUndo.TabIndex = 11;
            this.buttonDoUndo.Text = "<< Undo";
            this.buttonDoUndo.UseVisualStyleBackColor = true;
            this.buttonDoUndo.Click += new System.EventHandler(this.buttonDoUndo_Click);
            // 
            // buttonDoRedo
            // 
            this.buttonDoRedo.Enabled = false;
            this.buttonDoRedo.Location = new System.Drawing.Point(683, 1);
            this.buttonDoRedo.Name = "buttonDoRedo";
            this.buttonDoRedo.Size = new System.Drawing.Size(83, 30);
            this.buttonDoRedo.TabIndex = 12;
            this.buttonDoRedo.Text = "Redo >>";
            this.buttonDoRedo.UseVisualStyleBackColor = true;
            this.buttonDoRedo.Click += new System.EventHandler(this.buttonDoRedo_Click);
            // 
            // FormHost
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1053, 598);
            this.Controls.Add(this.buttonDoRedo);
            this.Controls.Add(this.buttonDoUndo);
            this.Controls.Add(this.buttonSimulateDatabaseAction);
            this.Controls.Add(this.buttonResetData);
            this.Controls.Add(this.buttonCancelChanges);
            this.Controls.Add(this.buttonSaveChanges);
            this.Controls.Add(this.tabControl1);
            this.Name = "FormHost";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NOToolsTests.CSharpTextEditor2";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private DataDisplayControl dataDisplayControl1;
        private DataDisplayControl dataDisplayControl2;
        private System.Windows.Forms.Button buttonResetData;
        private System.Windows.Forms.Button buttonCancelChanges;
        private System.Windows.Forms.Button buttonSaveChanges;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private DataDisplayControl dataDisplayControl3;
        private DataDisplayControl dataDisplayControl4;
        private System.Windows.Forms.Button buttonSimulateDatabaseAction;
        private System.Windows.Forms.Button buttonDoUndo;
        private System.Windows.Forms.Button buttonDoRedo;
    }
}

