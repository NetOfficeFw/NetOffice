namespace NetOffice.DeveloperToolbox
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
            this.checkBoxAccess = new System.Windows.Forms.CheckBox();
            this.checkBoxPowerPoint = new System.Windows.Forms.CheckBox();
            this.checkBoxOutlook = new System.Windows.Forms.CheckBox();
            this.checkBoxWord = new System.Windows.Forms.CheckBox();
            this.checkBoxExcel = new System.Windows.Forms.CheckBox();
            this.checkBoxProject = new System.Windows.Forms.CheckBox();
            this.checkBoxVisio = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // checkBoxAccess
            // 
            this.checkBoxAccess.AutoSize = true;
            this.checkBoxAccess.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxAccess.Location = new System.Drawing.Point(40, 130);
            this.checkBoxAccess.Name = "checkBoxAccess";
            this.checkBoxAccess.Size = new System.Drawing.Size(72, 20);
            this.checkBoxAccess.TabIndex = 9;
            this.checkBoxAccess.Text = "Access";
            this.checkBoxAccess.UseVisualStyleBackColor = true;
            this.checkBoxAccess.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxPowerPoint
            // 
            this.checkBoxPowerPoint.AutoSize = true;
            this.checkBoxPowerPoint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxPowerPoint.Location = new System.Drawing.Point(40, 107);
            this.checkBoxPowerPoint.Name = "checkBoxPowerPoint";
            this.checkBoxPowerPoint.Size = new System.Drawing.Size(98, 20);
            this.checkBoxPowerPoint.TabIndex = 8;
            this.checkBoxPowerPoint.Text = "Power Point";
            this.checkBoxPowerPoint.UseVisualStyleBackColor = true;
            this.checkBoxPowerPoint.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxOutlook
            // 
            this.checkBoxOutlook.AutoSize = true;
            this.checkBoxOutlook.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxOutlook.Location = new System.Drawing.Point(40, 82);
            this.checkBoxOutlook.Name = "checkBoxOutlook";
            this.checkBoxOutlook.Size = new System.Drawing.Size(73, 20);
            this.checkBoxOutlook.TabIndex = 7;
            this.checkBoxOutlook.Text = "Outlook";
            this.checkBoxOutlook.UseVisualStyleBackColor = true;
            this.checkBoxOutlook.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxWord
            // 
            this.checkBoxWord.AutoSize = true;
            this.checkBoxWord.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxWord.Location = new System.Drawing.Point(40, 57);
            this.checkBoxWord.Name = "checkBoxWord";
            this.checkBoxWord.Size = new System.Drawing.Size(60, 20);
            this.checkBoxWord.TabIndex = 6;
            this.checkBoxWord.Text = "Word";
            this.checkBoxWord.UseVisualStyleBackColor = true;
            this.checkBoxWord.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxExcel
            // 
            this.checkBoxExcel.AutoSize = true;
            this.checkBoxExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxExcel.Location = new System.Drawing.Point(40, 33);
            this.checkBoxExcel.Name = "checkBoxExcel";
            this.checkBoxExcel.Size = new System.Drawing.Size(60, 20);
            this.checkBoxExcel.TabIndex = 5;
            this.checkBoxExcel.Text = "Excel";
            this.checkBoxExcel.UseVisualStyleBackColor = true;
            this.checkBoxExcel.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxProject
            // 
            this.checkBoxProject.AutoSize = true;
            this.checkBoxProject.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxProject.Location = new System.Drawing.Point(40, 153);
            this.checkBoxProject.Name = "checkBoxProject";
            this.checkBoxProject.Size = new System.Drawing.Size(69, 20);
            this.checkBoxProject.TabIndex = 10;
            this.checkBoxProject.Text = "Project";
            this.checkBoxProject.UseVisualStyleBackColor = true;
            this.checkBoxProject.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxVisio
            // 
            this.checkBoxVisio.AutoSize = true;
            this.checkBoxVisio.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxVisio.Location = new System.Drawing.Point(40, 176);
            this.checkBoxVisio.Name = "checkBoxVisio";
            this.checkBoxVisio.Size = new System.Drawing.Size(57, 20);
            this.checkBoxVisio.TabIndex = 11;
            this.checkBoxVisio.Text = "Visio";
            this.checkBoxVisio.UseVisualStyleBackColor = true;
            this.checkBoxVisio.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // HostControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.checkBoxVisio);
            this.Controls.Add(this.checkBoxProject);
            this.Controls.Add(this.checkBoxAccess);
            this.Controls.Add(this.checkBoxPowerPoint);
            this.Controls.Add(this.checkBoxOutlook);
            this.Controls.Add(this.checkBoxWord);
            this.Controls.Add(this.checkBoxExcel);
            this.Name = "HostControl";
            this.Size = new System.Drawing.Size(611, 285);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.CheckBox checkBoxAccess;
        internal System.Windows.Forms.CheckBox checkBoxPowerPoint;
        internal System.Windows.Forms.CheckBox checkBoxOutlook;
        internal System.Windows.Forms.CheckBox checkBoxWord;
        internal System.Windows.Forms.CheckBox checkBoxExcel;
        internal System.Windows.Forms.CheckBox checkBoxProject;
        internal System.Windows.Forms.CheckBox checkBoxVisio;

    }
}
