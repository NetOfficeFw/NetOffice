namespace NOTools.CodeCommander.UI
{
    partial class InfoPane
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InfoPane));
            this.labelInfoHeader = new System.Windows.Forms.Label();
            this.linkLabelInfo = new System.Windows.Forms.LinkLabel();
            this.panelInfo = new System.Windows.Forms.Panel();
            this.richTextBoxInfo = new System.Windows.Forms.RichTextBox();
            this.checkBoxSaveSettings = new System.Windows.Forms.CheckBox();
            this.panelInfo.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelInfoHeader
            // 
            this.labelInfoHeader.AutoSize = true;
            this.labelInfoHeader.Location = new System.Drawing.Point(17, 361);
            this.labelInfoHeader.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelInfoHeader.Name = "labelInfoHeader";
            this.labelInfoHeader.Size = new System.Drawing.Size(76, 16);
            this.labelInfoHeader.TabIndex = 9;
            this.labelInfoHeader.Text = "Read more";
            // 
            // linkLabelInfo
            // 
            this.linkLabelInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelInfo.AutoSize = true;
            this.linkLabelInfo.Location = new System.Drawing.Point(93, 361);
            this.linkLabelInfo.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.linkLabelInfo.Name = "linkLabelInfo";
            this.linkLabelInfo.Size = new System.Drawing.Size(185, 16);
            this.linkLabelInfo.TabIndex = 8;
            this.linkLabelInfo.TabStop = true;
            this.linkLabelInfo.Tag = "http://netoffice.codeplex.com/wikipage?title=NO_Tools_DeveloperAddin";
            this.linkLabelInfo.Text = "netoffice.codeplex.com/Addin";
            this.linkLabelInfo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelInfo_LinkClicked);
            // 
            // panelInfo
            // 
            this.panelInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelInfo.Controls.Add(this.richTextBoxInfo);
            this.panelInfo.Location = new System.Drawing.Point(19, 46);
            this.panelInfo.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panelInfo.Name = "panelInfo";
            this.panelInfo.Size = new System.Drawing.Size(272, 303);
            this.panelInfo.TabIndex = 7;
            // 
            // richTextBoxInfo
            // 
            this.richTextBoxInfo.BackColor = System.Drawing.Color.LightYellow;
            this.richTextBoxInfo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBoxInfo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBoxInfo.Location = new System.Drawing.Point(0, 0);
            this.richTextBoxInfo.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.richTextBoxInfo.Name = "richTextBoxInfo";
            this.richTextBoxInfo.ReadOnly = true;
            this.richTextBoxInfo.Size = new System.Drawing.Size(270, 301);
            this.richTextBoxInfo.TabIndex = 1;
            this.richTextBoxInfo.Text = resources.GetString("richTextBoxInfo.Text");
            // 
            // checkBoxSaveSettings
            // 
            this.checkBoxSaveSettings.AutoSize = true;
            this.checkBoxSaveSettings.Checked = true;
            this.checkBoxSaveSettings.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSaveSettings.Location = new System.Drawing.Point(20, 16);
            this.checkBoxSaveSettings.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBoxSaveSettings.Name = "checkBoxSaveSettings";
            this.checkBoxSaveSettings.Size = new System.Drawing.Size(185, 20);
            this.checkBoxSaveSettings.TabIndex = 6;
            this.checkBoxSaveSettings.Text = "Save Settings (individualy)";
            this.checkBoxSaveSettings.UseVisualStyleBackColor = true;
            this.checkBoxSaveSettings.CheckedChanged += new System.EventHandler(this.checkBoxSaveSettings_CheckedChanged);
            // 
            // InfoPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.labelInfoHeader);
            this.Controls.Add(this.linkLabelInfo);
            this.Controls.Add(this.panelInfo);
            this.Controls.Add(this.checkBoxSaveSettings);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "InfoPane";
            this.Size = new System.Drawing.Size(312, 388);
            this.panelInfo.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelInfoHeader;
        private System.Windows.Forms.LinkLabel linkLabelInfo;
        private System.Windows.Forms.Panel panelInfo;
        private System.Windows.Forms.RichTextBox richTextBoxInfo;
        private System.Windows.Forms.CheckBox checkBoxSaveSettings;
    }
}
