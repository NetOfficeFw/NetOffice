namespace NetOffice.ProjectWizard
{
    partial class AddinGuiControl
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
            this.checkBoxRibbonUISupport = new System.Windows.Forms.CheckBox();
            this.checkBoxClassicUISupport = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // checkBoxRibbonUISupport
            // 
            this.checkBoxRibbonUISupport.AutoSize = true;
            this.checkBoxRibbonUISupport.Location = new System.Drawing.Point(30, 58);
            this.checkBoxRibbonUISupport.Name = "checkBoxRibbonUISupport";
            this.checkBoxRibbonUISupport.Size = new System.Drawing.Size(368, 17);
            this.checkBoxRibbonUISupport.TabIndex = 19;
            this.checkBoxRibbonUISupport.Text = "Ich möchte die Ribbon Oberfläche in neueren Office Versionen erweitern";
            this.checkBoxRibbonUISupport.UseVisualStyleBackColor = true;
            this.checkBoxRibbonUISupport.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxClassicUISupport
            // 
            this.checkBoxClassicUISupport.AutoSize = true;
            this.checkBoxClassicUISupport.Location = new System.Drawing.Point(30, 30);
            this.checkBoxClassicUISupport.Name = "checkBoxClassicUISupport";
            this.checkBoxClassicUISupport.Size = new System.Drawing.Size(364, 17);
            this.checkBoxClassicUISupport.TabIndex = 21;
            this.checkBoxClassicUISupport.Text = "Ich möchte die Benutzeroberfläche in älteren Office Versionen erweitern";
            this.checkBoxClassicUISupport.UseVisualStyleBackColor = true;
            this.checkBoxClassicUISupport.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // AddinGuiControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.checkBoxClassicUISupport);
            this.Controls.Add(this.checkBoxRibbonUISupport);
            this.Name = "AddinGuiControl";
            this.Size = new System.Drawing.Size(520, 200);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkBoxRibbonUISupport;
        private System.Windows.Forms.CheckBox checkBoxClassicUISupport;
    }
}
