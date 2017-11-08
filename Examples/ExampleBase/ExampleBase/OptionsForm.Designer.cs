namespace ExampleBase
{
    partial class OptionsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OptionsForm));
            this.radioButtonCommonFolder = new System.Windows.Forms.RadioButton();
            this.radioButtonApplicationFolder = new System.Windows.Forms.RadioButton();
            this.groupBoxFolder = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonDone = new System.Windows.Forms.Button();
            this.groupBoxFolder.SuspendLayout();
            this.SuspendLayout();
            // 
            // radioButtonCommonFolder
            // 
            this.radioButtonCommonFolder.AutoSize = true;
            this.radioButtonCommonFolder.Checked = true;
            this.radioButtonCommonFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonCommonFolder.Location = new System.Drawing.Point(28, 39);
            this.radioButtonCommonFolder.Name = "radioButtonCommonFolder";
            this.radioButtonCommonFolder.Size = new System.Drawing.Size(160, 20);
            this.radioButtonCommonFolder.TabIndex = 0;
            this.radioButtonCommonFolder.TabStop = true;
            this.radioButtonCommonFolder.Text = "Local Application Data";
            this.radioButtonCommonFolder.UseVisualStyleBackColor = true;
            // 
            // radioButtonApplicationFolder
            // 
            this.radioButtonApplicationFolder.AutoSize = true;
            this.radioButtonApplicationFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonApplicationFolder.Location = new System.Drawing.Point(28, 69);
            this.radioButtonApplicationFolder.Name = "radioButtonApplicationFolder";
            this.radioButtonApplicationFolder.Size = new System.Drawing.Size(163, 20);
            this.radioButtonApplicationFolder.TabIndex = 1;
            this.radioButtonApplicationFolder.Text = "This Application Folder";
            this.radioButtonApplicationFolder.UseVisualStyleBackColor = true;
            // 
            // groupBoxFolder
            // 
            this.groupBoxFolder.Controls.Add(this.label1);
            this.groupBoxFolder.Controls.Add(this.radioButtonApplicationFolder);
            this.groupBoxFolder.Controls.Add(this.radioButtonCommonFolder);
            this.groupBoxFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBoxFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.groupBoxFolder.Location = new System.Drawing.Point(23, 24);
            this.groupBoxFolder.Name = "groupBoxFolder";
            this.groupBoxFolder.Size = new System.Drawing.Size(365, 116);
            this.groupBoxFolder.TabIndex = 2;
            this.groupBoxFolder.TabStop = false;
            this.groupBoxFolder.Text = "Select base folder for generated documents";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(224, 71);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(142, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "(permissions required)";
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
            this.buttonDone.Location = new System.Drawing.Point(217, 146);
            this.buttonDone.Name = "buttonDone";
            this.buttonDone.Size = new System.Drawing.Size(171, 29);
            this.buttonDone.TabIndex = 4;
            this.buttonDone.Text = "Return to Examples";
            this.buttonDone.UseVisualStyleBackColor = true;
            this.buttonDone.Click += new System.EventHandler(this.buttonDone_Click);
            // 
            // OptionsForm
            // 
            this.AcceptButton = this.buttonDone;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.CancelButton = this.buttonDone;
            this.ClientSize = new System.Drawing.Size(425, 199);
            this.Controls.Add(this.buttonDone);
            this.Controls.Add(this.groupBoxFolder);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OptionsForm";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Options";
            this.groupBoxFolder.ResumeLayout(false);
            this.groupBoxFolder.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RadioButton radioButtonCommonFolder;
        private System.Windows.Forms.RadioButton radioButtonApplicationFolder;
        private System.Windows.Forms.GroupBox groupBoxFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonDone;

    }
}
