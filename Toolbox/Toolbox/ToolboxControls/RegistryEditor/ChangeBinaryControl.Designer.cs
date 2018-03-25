namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    partial class ChangeBinaryControl
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
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonAbort = new System.Windows.Forms.Button();
            this.textBoxName = new System.Windows.Forms.TextBox();
            this.labelName = new System.Windows.Forms.Label();
            this.hexBox = new NetOffice.DeveloperToolbox.Controls.Hex.HexBox();
            this.SuspendLayout();
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOK.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonOK.ForeColor = System.Drawing.Color.Black;
            this.buttonOK.Location = new System.Drawing.Point(216, 284);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(74, 22);
            this.buttonOK.TabIndex = 19;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonAbort
            // 
            this.buttonAbort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonAbort.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonAbort.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonAbort.ForeColor = System.Drawing.Color.Black;
            this.buttonAbort.Location = new System.Drawing.Point(303, 284);
            this.buttonAbort.Name = "buttonAbort";
            this.buttonAbort.Size = new System.Drawing.Size(74, 22);
            this.buttonAbort.TabIndex = 18;
            this.buttonAbort.Text = "Cancel";
            this.buttonAbort.UseVisualStyleBackColor = true;
            this.buttonAbort.Click += new System.EventHandler(this.buttonAbort_Click);
            // 
            // textBoxName
            // 
            this.textBoxName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxName.ForeColor = System.Drawing.Color.Black;
            this.textBoxName.Location = new System.Drawing.Point(20, 26);
            this.textBoxName.Name = "textBoxName";
            this.textBoxName.ReadOnly = true;
            this.textBoxName.Size = new System.Drawing.Size(357, 22);
            this.textBoxName.TabIndex = 17;
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.ForeColor = System.Drawing.Color.Black;
            this.labelName.Location = new System.Drawing.Point(17, 10);
            this.labelName.Name = "labelName";
            this.labelName.Size = new System.Drawing.Size(36, 13);
            this.labelName.TabIndex = 16;
            this.labelName.Text = "Name";
            // 
            // hexBox
            // 
            this.hexBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.hexBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.hexBox.BytesPerLine = 7;
            this.hexBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.hexBox.Font = new System.Drawing.Font("Lucida Console", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.hexBox.ForeColor = System.Drawing.Color.Black;
            this.hexBox.HexCasing = NetOffice.DeveloperToolbox.Controls.Hex.HexCasing.Lower;
            this.hexBox.LineInfoForeColor = System.Drawing.Color.Empty;
            this.hexBox.LineInfoVisible = true;
            this.hexBox.Location = new System.Drawing.Point(19, 57);
            this.hexBox.Name = "hexBox";
            this.hexBox.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            this.hexBox.ShadowSelectionColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(60)))), ((int)(((byte)(188)))), ((int)(((byte)(255)))));
            this.hexBox.Size = new System.Drawing.Size(357, 210);
            this.hexBox.StringViewVisible = true;
            this.hexBox.TabIndex = 20;
            this.hexBox.UseFixedBytesPerLine = true;
            this.hexBox.VScrollBarVisible = true;
            // 
            // ChangeBinaryControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.hexBox);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonAbort);
            this.Controls.Add(this.textBoxName);
            this.Controls.Add(this.labelName);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "ChangeBinaryControl";
            this.Size = new System.Drawing.Size(394, 314);
            this.Resize += new System.EventHandler(this.ChangeBinaryControl_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Controls.Hex.HexBox hexBox;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonAbort;
        private System.Windows.Forms.TextBox textBoxName;
        private System.Windows.Forms.Label labelName;
    }
}
