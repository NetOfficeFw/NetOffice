namespace NetOffice.DeveloperToolbox.Translation
{
    partial class ToolLanguageControl
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.languageSummaryControl1 = new NetOffice.DeveloperToolbox.Translation.LanguageSummaryControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.languageApplicationControl1 = new NetOffice.DeveloperToolbox.Translation.LanguageApplicationControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.languageComponentsControl1 = new NetOffice.DeveloperToolbox.Translation.LanguageComponentsControl();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(857, 550);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.languageSummaryControl1);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(849, 521);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Summary";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // languageSummaryControl1
            // 
            this.languageSummaryControl1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.languageSummaryControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.languageSummaryControl1.Location = new System.Drawing.Point(3, 3);
            this.languageSummaryControl1.Name = "languageSummaryControl1";
            this.languageSummaryControl1.Size = new System.Drawing.Size(843, 515);
            this.languageSummaryControl1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.languageApplicationControl1);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(849, 521);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Application";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // languageApplicationControl1
            // 
            this.languageApplicationControl1.BackColor = System.Drawing.SystemColors.Control;
            this.languageApplicationControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.languageApplicationControl1.Location = new System.Drawing.Point(3, 3);
            this.languageApplicationControl1.Name = "languageApplicationControl1";
            this.languageApplicationControl1.Size = new System.Drawing.Size(843, 515);
            this.languageApplicationControl1.TabIndex = 0;
            this.languageApplicationControl1.SelectionChanged += new System.EventHandler(this.languageApplicationControl1_SelectionChanged);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.languageComponentsControl1);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(849, 521);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Components";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // languageComponentsControl1
            // 
            this.languageComponentsControl1.BackColor = System.Drawing.SystemColors.Control;
            this.languageComponentsControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.languageComponentsControl1.Location = new System.Drawing.Point(0, 0);
            this.languageComponentsControl1.Name = "languageComponentsControl1";
            this.languageComponentsControl1.Size = new System.Drawing.Size(849, 521);
            this.languageComponentsControl1.TabIndex = 0;
            this.languageComponentsControl1.SelectionChanged += new System.EventHandler(this.languageComponentsControl1_SelectionChanged);
            // 
            // ToolLanguageControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Name = "ToolLanguageControl";
            this.Size = new System.Drawing.Size(857, 550);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private LanguageSummaryControl languageSummaryControl1;
        private System.Windows.Forms.TabPage tabPage3;
        private LanguageApplicationControl languageApplicationControl1;
        private LanguageComponentsControl languageComponentsControl1;
    }
}
