namespace ExampleBase
{
    partial class FormOptions
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormOptions));
            this.radioButtonCommonFolder = new System.Windows.Forms.RadioButton();
            this.radioButtonApplicationFolder = new System.Windows.Forms.RadioButton();
            this.groupBoxFolder = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBoxLanguage = new System.Windows.Forms.GroupBox();
            this.radioButtonLanguage1031 = new System.Windows.Forms.RadioButton();
            this.radioButtonLanguage1033 = new System.Windows.Forms.RadioButton();
            this.buttonDone = new System.Windows.Forms.Button();
            this.groupBoxFolder.SuspendLayout();
            this.groupBoxLanguage.SuspendLayout();
            this.SuspendLayout();
            // 
            // radioButtonCommonFolder
            // 
            this.radioButtonCommonFolder.AutoSize = true;
            this.radioButtonCommonFolder.Checked = true;
            this.radioButtonCommonFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonCommonFolder.Location = new System.Drawing.Point(19, 29);
            this.radioButtonCommonFolder.Name = "radioButtonCommonFolder";
            this.radioButtonCommonFolder.Size = new System.Drawing.Size(121, 17);
            this.radioButtonCommonFolder.TabIndex = 0;
            this.radioButtonCommonFolder.TabStop = true;
            this.radioButtonCommonFolder.Text = "Common Files Folder";
            this.radioButtonCommonFolder.UseVisualStyleBackColor = true;
            // 
            // radioButtonApplicationFolder
            // 
            this.radioButtonApplicationFolder.AutoSize = true;
            this.radioButtonApplicationFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonApplicationFolder.Location = new System.Drawing.Point(19, 52);
            this.radioButtonApplicationFolder.Name = "radioButtonApplicationFolder";
            this.radioButtonApplicationFolder.Size = new System.Drawing.Size(108, 17);
            this.radioButtonApplicationFolder.TabIndex = 1;
            this.radioButtonApplicationFolder.Text = "Application Folder";
            this.radioButtonApplicationFolder.UseVisualStyleBackColor = true;
            // 
            // groupBoxFolder
            // 
            this.groupBoxFolder.Controls.Add(this.label1);
            this.groupBoxFolder.Controls.Add(this.radioButtonApplicationFolder);
            this.groupBoxFolder.Controls.Add(this.radioButtonCommonFolder);
            this.groupBoxFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBoxFolder.Location = new System.Drawing.Point(23, 24);
            this.groupBoxFolder.Name = "groupBoxFolder";
            this.groupBoxFolder.Size = new System.Drawing.Size(282, 89);
            this.groupBoxFolder.TabIndex = 2;
            this.groupBoxFolder.TabStop = false;
            this.groupBoxFolder.Text = "Select base folder for generated documents";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(134, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "(permissions required)";
            // 
            // groupBoxLanguage
            // 
            this.groupBoxLanguage.Controls.Add(this.radioButtonLanguage1031);
            this.groupBoxLanguage.Controls.Add(this.radioButtonLanguage1033);
            this.groupBoxLanguage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBoxLanguage.Location = new System.Drawing.Point(23, 130);
            this.groupBoxLanguage.Name = "groupBoxLanguage";
            this.groupBoxLanguage.Size = new System.Drawing.Size(282, 89);
            this.groupBoxLanguage.TabIndex = 3;
            this.groupBoxLanguage.TabStop = false;
            this.groupBoxLanguage.Text = "Select Language";
            // 
            // radioButtonLanguage1031
            // 
            this.radioButtonLanguage1031.AutoSize = true;
            this.radioButtonLanguage1031.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonLanguage1031.Location = new System.Drawing.Point(19, 53);
            this.radioButtonLanguage1031.Name = "radioButtonLanguage1031";
            this.radioButtonLanguage1031.Size = new System.Drawing.Size(110, 17);
            this.radioButtonLanguage1031.TabIndex = 3;
            this.radioButtonLanguage1031.Text = "German (Deutsch)";
            this.radioButtonLanguage1031.UseVisualStyleBackColor = true;
            // 
            // radioButtonLanguage1033
            // 
            this.radioButtonLanguage1033.AutoSize = true;
            this.radioButtonLanguage1033.Checked = true;
            this.radioButtonLanguage1033.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonLanguage1033.Location = new System.Drawing.Point(19, 30);
            this.radioButtonLanguage1033.Name = "radioButtonLanguage1033";
            this.radioButtonLanguage1033.Size = new System.Drawing.Size(82, 17);
            this.radioButtonLanguage1033.TabIndex = 2;
            this.radioButtonLanguage1033.TabStop = true;
            this.radioButtonLanguage1033.Text = "English (US)";
            this.radioButtonLanguage1033.UseVisualStyleBackColor = true;
            this.radioButtonLanguage1033.CheckedChanged += new System.EventHandler(this.radioButtonLanguage1033_CheckedChanged);
            // 
            // buttonDone
            // 
            this.buttonDone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDone.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonDone.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonDone.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonDone.ForeColor = System.Drawing.Color.Blue;
            this.buttonDone.Image = ((System.Drawing.Image)(resources.GetObject("buttonDone.Image")));
            this.buttonDone.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonDone.Location = new System.Drawing.Point(134, 237);
            this.buttonDone.Name = "buttonDone";
            this.buttonDone.Size = new System.Drawing.Size(171, 29);
            this.buttonDone.TabIndex = 4;
            this.buttonDone.Text = "Return to Examples";
            this.buttonDone.UseVisualStyleBackColor = true;
            this.buttonDone.Click += new System.EventHandler(this.buttonDone_Click);
            // 
            // FormOptions
            // 
            this.AcceptButton = this.buttonDone;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.CancelButton = this.buttonDone;
            this.ClientSize = new System.Drawing.Size(333, 283);
            this.Controls.Add(this.buttonDone);
            this.Controls.Add(this.groupBoxLanguage);
            this.Controls.Add(this.groupBoxFolder);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormOptions";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Options";
            this.groupBoxFolder.ResumeLayout(false);
            this.groupBoxFolder.PerformLayout();
            this.groupBoxLanguage.ResumeLayout(false);
            this.groupBoxLanguage.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RadioButton radioButtonCommonFolder;
        private System.Windows.Forms.RadioButton radioButtonApplicationFolder;
        private System.Windows.Forms.GroupBox groupBoxFolder;
        private System.Windows.Forms.GroupBox groupBoxLanguage;
        private System.Windows.Forms.RadioButton radioButtonLanguage1031;
        private System.Windows.Forms.RadioButton radioButtonLanguage1033;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonDone;

    }
}
