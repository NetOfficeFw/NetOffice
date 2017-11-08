namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    partial class SummaryControl
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
            this.labelSummaryValue = new System.Windows.Forms.Label();
            this.labelSummaryHeader = new System.Windows.Forms.Label();
            this.labelSummaryCaption = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelSummaryValue
            // 
            this.labelSummaryValue.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelSummaryValue.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSummaryValue.ForeColor = System.Drawing.Color.Black;
            this.labelSummaryValue.Location = new System.Drawing.Point(296, 61);
            this.labelSummaryValue.Name = "labelSummaryValue";
            this.labelSummaryValue.Size = new System.Drawing.Size(391, 217);
            this.labelSummaryValue.TabIndex = 19;
            // 
            // labelSummaryHeader
            // 
            this.labelSummaryHeader.AutoSize = true;
            this.labelSummaryHeader.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSummaryHeader.ForeColor = System.Drawing.Color.Black;
            this.labelSummaryHeader.Location = new System.Drawing.Point(42, 34);
            this.labelSummaryHeader.Name = "labelSummaryHeader";
            this.labelSummaryHeader.Size = new System.Drawing.Size(108, 16);
            this.labelSummaryHeader.TabIndex = 18;
            this.labelSummaryHeader.Text = "Summary Table";
            // 
            // labelSummaryCaption
            // 
            this.labelSummaryCaption.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.labelSummaryCaption.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSummaryCaption.ForeColor = System.Drawing.Color.Black;
            this.labelSummaryCaption.Location = new System.Drawing.Point(42, 61);
            this.labelSummaryCaption.Name = "labelSummaryCaption";
            this.labelSummaryCaption.Size = new System.Drawing.Size(246, 217);
            this.labelSummaryCaption.TabIndex = 17;
            // 
            // SummaryControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.labelSummaryValue);
            this.Controls.Add(this.labelSummaryHeader);
            this.Controls.Add(this.labelSummaryCaption);
            this.Name = "SummaryControl";
            this.Size = new System.Drawing.Size(744, 311);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelSummaryValue;
        private System.Windows.Forms.Label labelSummaryHeader;
        private System.Windows.Forms.Label labelSummaryCaption;
    }
}
