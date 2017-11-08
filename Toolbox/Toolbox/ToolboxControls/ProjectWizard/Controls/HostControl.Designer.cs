namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    partial class HostControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HostControl));
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.labelHint = new System.Windows.Forms.Label();
            this.checkBoxVisio = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxProject = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxAccess = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxPowerPoint = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxOutlook = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxWord = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxExcel = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(185, 38);
            this.pictureBox2.Margin = new System.Windows.Forms.Padding(4);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(18, 18);
            this.pictureBox2.TabIndex = 118;
            this.pictureBox2.TabStop = false;
            // 
            // labelHint
            // 
            this.labelHint.AutoSize = true;
            this.labelHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelHint.ForeColor = System.Drawing.Color.DimGray;
            this.labelHint.Location = new System.Drawing.Point(211, 38);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(472, 16);
            this.labelHint.TabIndex = 119;
            this.labelHint.Text = "Use also number keys(1-7) on your keyboard to select/deselect an application";
            // 
            // checkBoxVisio
            // 
            this.checkBoxVisio.AutoSize = true;
            this.checkBoxVisio.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxVisio.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxVisio.ForeColor = System.Drawing.Color.Black;
            this.checkBoxVisio.Location = new System.Drawing.Point(40, 176);
            this.checkBoxVisio.Name = "checkBoxVisio";
            this.checkBoxVisio.Size = new System.Drawing.Size(52, 21);
            this.checkBoxVisio.TabIndex = 11;
            this.checkBoxVisio.Text = "Visio";
            this.checkBoxVisio.UseVisualStyleBackColor = true;
            this.checkBoxVisio.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxProject
            // 
            this.checkBoxProject.AutoSize = true;
            this.checkBoxProject.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxProject.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxProject.ForeColor = System.Drawing.Color.Black;
            this.checkBoxProject.Location = new System.Drawing.Point(40, 153);
            this.checkBoxProject.Name = "checkBoxProject";
            this.checkBoxProject.Size = new System.Drawing.Size(64, 21);
            this.checkBoxProject.TabIndex = 10;
            this.checkBoxProject.Text = "Project";
            this.checkBoxProject.UseVisualStyleBackColor = true;
            this.checkBoxProject.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxAccess
            // 
            this.checkBoxAccess.AutoSize = true;
            this.checkBoxAccess.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxAccess.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxAccess.ForeColor = System.Drawing.Color.Black;
            this.checkBoxAccess.Location = new System.Drawing.Point(40, 130);
            this.checkBoxAccess.Name = "checkBoxAccess";
            this.checkBoxAccess.Size = new System.Drawing.Size(63, 21);
            this.checkBoxAccess.TabIndex = 9;
            this.checkBoxAccess.Text = "Access";
            this.checkBoxAccess.UseVisualStyleBackColor = true;
            this.checkBoxAccess.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxPowerPoint
            // 
            this.checkBoxPowerPoint.AutoSize = true;
            this.checkBoxPowerPoint.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxPowerPoint.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxPowerPoint.ForeColor = System.Drawing.Color.Black;
            this.checkBoxPowerPoint.Location = new System.Drawing.Point(40, 107);
            this.checkBoxPowerPoint.Name = "checkBoxPowerPoint";
            this.checkBoxPowerPoint.Size = new System.Drawing.Size(93, 21);
            this.checkBoxPowerPoint.TabIndex = 8;
            this.checkBoxPowerPoint.Text = "Power Point";
            this.checkBoxPowerPoint.UseVisualStyleBackColor = true;
            this.checkBoxPowerPoint.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxOutlook
            // 
            this.checkBoxOutlook.AutoSize = true;
            this.checkBoxOutlook.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxOutlook.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxOutlook.ForeColor = System.Drawing.Color.Black;
            this.checkBoxOutlook.Location = new System.Drawing.Point(40, 82);
            this.checkBoxOutlook.Name = "checkBoxOutlook";
            this.checkBoxOutlook.Size = new System.Drawing.Size(70, 21);
            this.checkBoxOutlook.TabIndex = 7;
            this.checkBoxOutlook.Text = "Outlook";
            this.checkBoxOutlook.UseVisualStyleBackColor = true;
            this.checkBoxOutlook.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxWord
            // 
            this.checkBoxWord.AutoSize = true;
            this.checkBoxWord.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxWord.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxWord.ForeColor = System.Drawing.Color.Black;
            this.checkBoxWord.Location = new System.Drawing.Point(40, 57);
            this.checkBoxWord.Name = "checkBoxWord";
            this.checkBoxWord.Size = new System.Drawing.Size(57, 21);
            this.checkBoxWord.TabIndex = 6;
            this.checkBoxWord.Text = "Word";
            this.checkBoxWord.UseVisualStyleBackColor = true;
            this.checkBoxWord.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxExcel
            // 
            this.checkBoxExcel.AutoSize = true;
            this.checkBoxExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxExcel.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxExcel.ForeColor = System.Drawing.Color.Black;
            this.checkBoxExcel.Location = new System.Drawing.Point(40, 33);
            this.checkBoxExcel.Name = "checkBoxExcel";
            this.checkBoxExcel.Size = new System.Drawing.Size(53, 21);
            this.checkBoxExcel.TabIndex = 5;
            this.checkBoxExcel.Text = "Excel";
            this.checkBoxExcel.UseVisualStyleBackColor = true;
            this.checkBoxExcel.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // HostControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.labelHint);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.checkBoxVisio);
            this.Controls.Add(this.checkBoxProject);
            this.Controls.Add(this.checkBoxAccess);
            this.Controls.Add(this.checkBoxPowerPoint);
            this.Controls.Add(this.checkBoxOutlook);
            this.Controls.Add(this.checkBoxWord);
            this.Controls.Add(this.checkBoxExcel);
            this.Name = "HostControl";
            this.Size = new System.Drawing.Size(659, 285);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxAccess;
        internal NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxPowerPoint;
        internal NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxOutlook;
        internal NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxWord;
        internal NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxExcel;
        internal NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxProject;
        internal NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxVisio;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label labelHint;

    }
}
