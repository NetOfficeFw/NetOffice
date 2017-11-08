namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GuiControl));
            this.checkBoxClassicUISupport = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxRibbonUISupport = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxTaskPaneSupport = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.checkBoxToogleButton = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            this.labelHint = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // checkBoxClassicUISupport
            // 
            this.checkBoxClassicUISupport.AutoSize = true;
            this.checkBoxClassicUISupport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxClassicUISupport.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxClassicUISupport.ForeColor = System.Drawing.Color.Black;
            this.checkBoxClassicUISupport.Location = new System.Drawing.Point(40, 33);
            this.checkBoxClassicUISupport.Name = "checkBoxClassicUISupport";
            this.checkBoxClassicUISupport.Size = new System.Drawing.Size(237, 21);
            this.checkBoxClassicUISupport.TabIndex = 23;
            this.checkBoxClassicUISupport.Text = "I want customize the classic Office UI";
            this.checkBoxClassicUISupport.UseVisualStyleBackColor = true;
            this.checkBoxClassicUISupport.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxRibbonUISupport
            // 
            this.checkBoxRibbonUISupport.AutoSize = true;
            this.checkBoxRibbonUISupport.Checked = true;
            this.checkBoxRibbonUISupport.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxRibbonUISupport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxRibbonUISupport.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxRibbonUISupport.ForeColor = System.Drawing.Color.Black;
            this.checkBoxRibbonUISupport.Location = new System.Drawing.Point(40, 61);
            this.checkBoxRibbonUISupport.Name = "checkBoxRibbonUISupport";
            this.checkBoxRibbonUISupport.Size = new System.Drawing.Size(204, 21);
            this.checkBoxRibbonUISupport.TabIndex = 22;
            this.checkBoxRibbonUISupport.Text = "I want customize the Ribbon UI";
            this.checkBoxRibbonUISupport.UseVisualStyleBackColor = true;
            this.checkBoxRibbonUISupport.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // checkBoxTaskPaneSupport
            // 
            this.checkBoxTaskPaneSupport.AutoSize = true;
            this.checkBoxTaskPaneSupport.Checked = true;
            this.checkBoxTaskPaneSupport.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxTaskPaneSupport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxTaskPaneSupport.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxTaskPaneSupport.ForeColor = System.Drawing.Color.Black;
            this.checkBoxTaskPaneSupport.Location = new System.Drawing.Point(40, 89);
            this.checkBoxTaskPaneSupport.Name = "checkBoxTaskPaneSupport";
            this.checkBoxTaskPaneSupport.Size = new System.Drawing.Size(190, 21);
            this.checkBoxTaskPaneSupport.TabIndex = 24;
            this.checkBoxTaskPaneSupport.Text = "Ich want a custom Task Pane";
            this.checkBoxTaskPaneSupport.UseVisualStyleBackColor = true;
            this.checkBoxTaskPaneSupport.CheckedChanged += new System.EventHandler(this.checkBoxTaskPaneSupport_CheckedChanged);
            // 
            // checkBoxToogleButton
            // 
            this.checkBoxToogleButton.AutoSize = true;
            this.checkBoxToogleButton.Checked = true;
            this.checkBoxToogleButton.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxToogleButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxToogleButton.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxToogleButton.ForeColor = System.Drawing.Color.Black;
            this.checkBoxToogleButton.Location = new System.Drawing.Point(40, 117);
            this.checkBoxToogleButton.Name = "checkBoxToogleButton";
            this.checkBoxToogleButton.Size = new System.Drawing.Size(323, 21);
            this.checkBoxToogleButton.TabIndex = 25;
            this.checkBoxToogleButton.Text = "Create a Toogle Button for the Task Pane (Visibility)";
            this.checkBoxToogleButton.UseVisualStyleBackColor = true;
            this.checkBoxToogleButton.Visible = false;
            this.checkBoxToogleButton.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // labelHint
            // 
            this.labelHint.AutoSize = true;
            this.labelHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelHint.ForeColor = System.Drawing.Color.DimGray;
            this.labelHint.Location = new System.Drawing.Point(63, 174);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(443, 16);
            this.labelHint.TabIndex = 121;
            this.labelHint.Text = "Use also number keys(1-4) on your keyboard to select/deselect an option";
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(37, 174);
            this.pictureBox2.Margin = new System.Windows.Forms.Padding(4);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(18, 18);
            this.pictureBox2.TabIndex = 120;
            this.pictureBox2.TabStop = false;
            // 
            // GuiControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.labelHint);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.checkBoxToogleButton);
            this.Controls.Add(this.checkBoxTaskPaneSupport);
            this.Controls.Add(this.checkBoxClassicUISupport);
            this.Controls.Add(this.checkBoxRibbonUISupport);
            this.Name = "GuiControl";
            this.Size = new System.Drawing.Size(611, 285);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxClassicUISupport;
        private NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxRibbonUISupport;
        private NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxTaskPaneSupport;
        private NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxToogleButton;
        private System.Windows.Forms.Label labelHint;
        private System.Windows.Forms.PictureBox pictureBox2;
    }
}
