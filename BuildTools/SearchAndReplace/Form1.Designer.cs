namespace NOBuildTools.SearchAndReplace
{
    partial class Form1
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

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.TextBoxFolder = new System.Windows.Forms.TextBox();
            this.ButttonChooseFolder = new System.Windows.Forms.Button();
            this.RichTextBoxLog = new System.Windows.Forms.RichTextBox();
            this.LabelFolder = new System.Windows.Forms.Label();
            this.ButtonStart = new System.Windows.Forms.Button();
            this.TextBoxFilter = new System.Windows.Forms.TextBox();
            this.LabelFilter = new System.Windows.Forms.Label();
            this.LabelSearch = new System.Windows.Forms.Label();
            this.TextBoxSearch = new System.Windows.Forms.TextBox();
            this.TextBoxReplace = new System.Windows.Forms.TextBox();
            this.LabelReplace = new System.Windows.Forms.Label();
            this.ButtonLoadConfig = new System.Windows.Forms.Button();
            this.ButtonSaveConfig = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TextBoxFolder
            // 
            this.TextBoxFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxFolder.Location = new System.Drawing.Point(67, 21);
            this.TextBoxFolder.Name = "TextBoxFolder";
            this.TextBoxFolder.Size = new System.Drawing.Size(431, 20);
            this.TextBoxFolder.TabIndex = 0;
            // 
            // ButttonChooseFolder
            // 
            this.ButttonChooseFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ButttonChooseFolder.Location = new System.Drawing.Point(507, 20);
            this.ButttonChooseFolder.Name = "ButttonChooseFolder";
            this.ButttonChooseFolder.Size = new System.Drawing.Size(40, 22);
            this.ButttonChooseFolder.TabIndex = 1;
            this.ButttonChooseFolder.Text = "...";
            this.ButttonChooseFolder.UseVisualStyleBackColor = true;
            this.ButttonChooseFolder.Click += new System.EventHandler(this.ButtonChooseFolder_Click);
            // 
            // RichTextBoxLog
            // 
            this.RichTextBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.RichTextBoxLog.Location = new System.Drawing.Point(67, 209);
            this.RichTextBoxLog.Name = "RichTextBoxLog";
            this.RichTextBoxLog.ReadOnly = true;
            this.RichTextBoxLog.Size = new System.Drawing.Size(746, 282);
            this.RichTextBoxLog.TabIndex = 2;
            this.RichTextBoxLog.Text = "";
            // 
            // LabelFolder
            // 
            this.LabelFolder.AutoSize = true;
            this.LabelFolder.Location = new System.Drawing.Point(12, 24);
            this.LabelFolder.Name = "LabelFolder";
            this.LabelFolder.Size = new System.Drawing.Size(49, 13);
            this.LabelFolder.TabIndex = 3;
            this.LabelFolder.Text = "Directory";
            // 
            // ButtonStart
            // 
            this.ButtonStart.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.ButtonStart.Location = new System.Drawing.Point(67, 159);
            this.ButtonStart.Name = "ButtonStart";
            this.ButtonStart.Size = new System.Drawing.Size(746, 23);
            this.ButtonStart.TabIndex = 4;
            this.ButtonStart.Text = "Start";
            this.ButtonStart.UseVisualStyleBackColor = true;
            this.ButtonStart.Click += new System.EventHandler(this.ButtonStart_Click);
            // 
            // TextBoxFilter
            // 
            this.TextBoxFilter.Location = new System.Drawing.Point(67, 62);
            this.TextBoxFilter.Multiline = true;
            this.TextBoxFilter.Name = "TextBoxFilter";
            this.TextBoxFilter.Size = new System.Drawing.Size(141, 79);
            this.TextBoxFilter.TabIndex = 5;
            this.TextBoxFilter.Text = "*.csproj\r\n*.vbproj";
            // 
            // LabelFilter
            // 
            this.LabelFilter.AutoSize = true;
            this.LabelFilter.Location = new System.Drawing.Point(12, 64);
            this.LabelFilter.Name = "LabelFilter";
            this.LabelFilter.Size = new System.Drawing.Size(29, 13);
            this.LabelFilter.TabIndex = 6;
            this.LabelFilter.Text = "Filter";
            // 
            // LabelSearch
            // 
            this.LabelSearch.AutoSize = true;
            this.LabelSearch.Location = new System.Drawing.Point(214, 63);
            this.LabelSearch.Name = "LabelSearch";
            this.LabelSearch.Size = new System.Drawing.Size(41, 13);
            this.LabelSearch.TabIndex = 7;
            this.LabelSearch.Text = "Search";
            // 
            // TextBoxSearch
            // 
            this.TextBoxSearch.Location = new System.Drawing.Point(261, 62);
            this.TextBoxSearch.Multiline = true;
            this.TextBoxSearch.Name = "TextBoxSearch";
            this.TextBoxSearch.Size = new System.Drawing.Size(237, 79);
            this.TextBoxSearch.TabIndex = 8;
            this.TextBoxSearch.Text = "<EmbedInteropTypes>True";
            // 
            // TextBoxReplace
            // 
            this.TextBoxReplace.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxReplace.Location = new System.Drawing.Point(557, 62);
            this.TextBoxReplace.Multiline = true;
            this.TextBoxReplace.Name = "TextBoxReplace";
            this.TextBoxReplace.Size = new System.Drawing.Size(256, 79);
            this.TextBoxReplace.TabIndex = 10;
            this.TextBoxReplace.Text = "<EmbedInteropTypes>False";
            // 
            // LabelReplace
            // 
            this.LabelReplace.AutoSize = true;
            this.LabelReplace.Location = new System.Drawing.Point(504, 63);
            this.LabelReplace.Name = "LabelReplace";
            this.LabelReplace.Size = new System.Drawing.Size(47, 13);
            this.LabelReplace.TabIndex = 9;
            this.LabelReplace.Text = "Replace";
            // 
            // ButtonLoadConfig
            // 
            this.ButtonLoadConfig.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ButtonLoadConfig.Location = new System.Drawing.Point(558, 18);
            this.ButtonLoadConfig.Name = "ButtonLoadConfig";
            this.ButtonLoadConfig.Size = new System.Drawing.Size(127, 27);
            this.ButtonLoadConfig.TabIndex = 11;
            this.ButtonLoadConfig.Text = "Load Config";
            this.ButtonLoadConfig.UseVisualStyleBackColor = true;
            this.ButtonLoadConfig.Click += new System.EventHandler(this.ButtonLoadConfig_Click);
            // 
            // ButtonSaveConfig
            // 
            this.ButtonSaveConfig.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ButtonSaveConfig.Location = new System.Drawing.Point(697, 18);
            this.ButtonSaveConfig.Name = "ButtonSaveConfig";
            this.ButtonSaveConfig.Size = new System.Drawing.Size(116, 27);
            this.ButtonSaveConfig.TabIndex = 12;
            this.ButtonSaveConfig.Text = "Save Config";
            this.ButtonSaveConfig.UseVisualStyleBackColor = true;
            this.ButtonSaveConfig.Click += new System.EventHandler(this.ButtonSaveConfig_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(839, 512);
            this.Controls.Add(this.ButtonSaveConfig);
            this.Controls.Add(this.ButtonLoadConfig);
            this.Controls.Add(this.TextBoxReplace);
            this.Controls.Add(this.LabelReplace);
            this.Controls.Add(this.TextBoxSearch);
            this.Controls.Add(this.LabelSearch);
            this.Controls.Add(this.LabelFilter);
            this.Controls.Add(this.TextBoxFilter);
            this.Controls.Add(this.ButtonStart);
            this.Controls.Add(this.LabelFolder);
            this.Controls.Add(this.RichTextBoxLog);
            this.Controls.Add(this.ButttonChooseFolder);
            this.Controls.Add(this.TextBoxFolder);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NOBuildTools.SearchAndReplace";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TextBoxFolder;
        private System.Windows.Forms.Button ButttonChooseFolder;
        private System.Windows.Forms.RichTextBox RichTextBoxLog;
        private System.Windows.Forms.Label LabelFolder;
        private System.Windows.Forms.Button ButtonStart;
        private System.Windows.Forms.TextBox TextBoxFilter;
        private System.Windows.Forms.Label LabelFilter;
        private System.Windows.Forms.Label LabelSearch;
        private System.Windows.Forms.TextBox TextBoxSearch;
        private System.Windows.Forms.TextBox TextBoxReplace;
        private System.Windows.Forms.Label LabelReplace;
        private System.Windows.Forms.Button ButtonLoadConfig;
        private System.Windows.Forms.Button ButtonSaveConfig;
    }
}

