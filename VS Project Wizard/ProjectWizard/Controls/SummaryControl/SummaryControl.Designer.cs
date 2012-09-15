namespace NetOffice.ProjectWizard
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
            this.labelSummaryCaption = new System.Windows.Forms.Label();
            this.labelSummaryHeader = new System.Windows.Forms.Label();
            this.labelSummaryValue = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelSummaryCaption
            // 
            this.labelSummaryCaption.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.labelSummaryCaption.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSummaryCaption.Location = new System.Drawing.Point(21, 41);
            this.labelSummaryCaption.Name = "labelSummaryCaption";
            this.labelSummaryCaption.Size = new System.Drawing.Size(182, 155);
            this.labelSummaryCaption.TabIndex = 0;
            // 
            // labelSummaryHeader
            // 
            this.labelSummaryHeader.AutoSize = true;
            this.labelSummaryHeader.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSummaryHeader.Location = new System.Drawing.Point(17, 20);
            this.labelSummaryHeader.Name = "labelSummaryHeader";
            this.labelSummaryHeader.Size = new System.Drawing.Size(246, 15);
            this.labelSummaryHeader.TabIndex = 15;
            this.labelSummaryHeader.Text = "Ausgwählte Einstellungen in der Übersicht";
            // 
            // labelSummaryValue
            // 
            this.labelSummaryValue.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelSummaryValue.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSummaryValue.Location = new System.Drawing.Point(209, 41);
            this.labelSummaryValue.Name = "labelSummaryValue";
            this.labelSummaryValue.Size = new System.Drawing.Size(303, 155);
            this.labelSummaryValue.TabIndex = 16;
            // 
            // SummaryControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelSummaryValue);
            this.Controls.Add(this.labelSummaryHeader);
            this.Controls.Add(this.labelSummaryCaption);
            this.Name = "SummaryControl";
            this.Size = new System.Drawing.Size(524, 212);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelSummaryCaption;
        private System.Windows.Forms.Label labelSummaryHeader;
        private System.Windows.Forms.Label labelSummaryValue;
    }
}
