namespace NetOffice.DeveloperToolbox
{
    partial class GuiControl
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
            this.checkBoxClassicUISupport = new System.Windows.Forms.CheckBox();
            this.checkBoxRibbonUISupport = new System.Windows.Forms.CheckBox();
            this.checkBoxTaskPaneSupport = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // checkBoxClassicUISupport
            // 
            this.checkBoxClassicUISupport.AutoSize = true;
            this.checkBoxClassicUISupport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxClassicUISupport.Location = new System.Drawing.Point(40, 33);
            this.checkBoxClassicUISupport.Name = "checkBoxClassicUISupport";
            this.checkBoxClassicUISupport.Size = new System.Drawing.Size(447, 20);
            this.checkBoxClassicUISupport.TabIndex = 23;
            this.checkBoxClassicUISupport.Text = "Ich möchte die Benutzeroberfläche in älteren Office Versionen erweitern";
            this.checkBoxClassicUISupport.UseVisualStyleBackColor = true;
            this.checkBoxClassicUISupport.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxRibbonUISupport
            // 
            this.checkBoxRibbonUISupport.AutoSize = true;
            this.checkBoxRibbonUISupport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxRibbonUISupport.Location = new System.Drawing.Point(40, 61);
            this.checkBoxRibbonUISupport.Name = "checkBoxRibbonUISupport";
            this.checkBoxRibbonUISupport.Size = new System.Drawing.Size(452, 20);
            this.checkBoxRibbonUISupport.TabIndex = 22;
            this.checkBoxRibbonUISupport.Text = "Ich möchte die Ribbon Oberfläche in neueren Office Versionen erweitern";
            this.checkBoxRibbonUISupport.UseVisualStyleBackColor = true;
            this.checkBoxRibbonUISupport.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxTaskPaneSupport
            // 
            this.checkBoxTaskPaneSupport.AutoSize = true;
            this.checkBoxTaskPaneSupport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxTaskPaneSupport.Location = new System.Drawing.Point(40, 89);
            this.checkBoxTaskPaneSupport.Name = "checkBoxTaskPaneSupport";
            this.checkBoxTaskPaneSupport.Size = new System.Drawing.Size(315, 20);
            this.checkBoxTaskPaneSupport.TabIndex = 24;
            this.checkBoxTaskPaneSupport.Text = "Ich möchte eine Task Pane zur Verfügung stellen";
            this.checkBoxTaskPaneSupport.UseVisualStyleBackColor = true;
            // 
            // GuiControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.checkBoxTaskPaneSupport);
            this.Controls.Add(this.checkBoxClassicUISupport);
            this.Controls.Add(this.checkBoxRibbonUISupport);
            this.Name = "GuiControl";
            this.Size = new System.Drawing.Size(611, 285);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkBoxClassicUISupport;
        private System.Windows.Forms.CheckBox checkBoxRibbonUISupport;
        private System.Windows.Forms.CheckBox checkBoxTaskPaneSupport;
    }
}
