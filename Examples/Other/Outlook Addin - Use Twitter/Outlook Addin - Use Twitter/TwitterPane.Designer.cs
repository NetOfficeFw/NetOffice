namespace Sample.Addin
{
    partial class TwitterPane
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TwitterPane));
            this.buttonAddinSettings = new System.Windows.Forms.Button();
            this.buttonMain = new System.Windows.Forms.Button();
            this.splitContainerButtons = new System.Windows.Forms.SplitContainer();
            this.errorPane = new Sample.Addin.ErrorPane();
            this.settingsPane = new Sample.Addin.SettingsPane();
            this.tweetGrid = new Sample.Addin.TweetGrid();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerButtons)).BeginInit();
            this.splitContainerButtons.Panel1.SuspendLayout();
            this.splitContainerButtons.Panel2.SuspendLayout();
            this.splitContainerButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonAddinSettings
            // 
            this.buttonAddinSettings.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonAddinSettings.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonAddinSettings.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonAddinSettings.Image = ((System.Drawing.Image)(resources.GetObject("buttonAddinSettings.Image")));
            this.buttonAddinSettings.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonAddinSettings.Location = new System.Drawing.Point(0, 0);
            this.buttonAddinSettings.Margin = new System.Windows.Forms.Padding(4);
            this.buttonAddinSettings.Name = "buttonAddinSettings";
            this.buttonAddinSettings.Size = new System.Drawing.Size(151, 30);
            this.buttonAddinSettings.TabIndex = 5;
            this.buttonAddinSettings.Text = "   Settings";
            this.buttonAddinSettings.UseVisualStyleBackColor = true;
            this.buttonAddinSettings.Click += new System.EventHandler(this.button_Click);
            // 
            // buttonMain
            // 
            this.buttonMain.BackColor = System.Drawing.Color.Goldenrod;
            this.buttonMain.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonMain.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonMain.Image = ((System.Drawing.Image)(resources.GetObject("buttonMain.Image")));
            this.buttonMain.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonMain.Location = new System.Drawing.Point(0, 0);
            this.buttonMain.Margin = new System.Windows.Forms.Padding(4);
            this.buttonMain.Name = "buttonMain";
            this.buttonMain.Size = new System.Drawing.Size(142, 30);
            this.buttonMain.TabIndex = 6;
            this.buttonMain.Text = "Main Page";
            this.buttonMain.UseVisualStyleBackColor = false;
            this.buttonMain.Click += new System.EventHandler(this.button_Click);
            // 
            // splitContainerButtons
            // 
            this.splitContainerButtons.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainerButtons.Location = new System.Drawing.Point(3, 702);
            this.splitContainerButtons.Name = "splitContainerButtons";
            // 
            // splitContainerButtons.Panel1
            // 
            this.splitContainerButtons.Panel1.Controls.Add(this.buttonMain);
            // 
            // splitContainerButtons.Panel2
            // 
            this.splitContainerButtons.Panel2.Controls.Add(this.buttonAddinSettings);
            this.splitContainerButtons.Size = new System.Drawing.Size(297, 30);
            this.splitContainerButtons.SplitterDistance = 142;
            this.splitContainerButtons.TabIndex = 11;
            // 
            // errorPane
            // 
            this.errorPane.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.errorPane.BackColor = System.Drawing.Color.LightSteelBlue;
            this.errorPane.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.errorPane.Location = new System.Drawing.Point(3, 672);
            this.errorPane.Margin = new System.Windows.Forms.Padding(4);
            this.errorPane.Name = "errorPane";
            this.errorPane.Size = new System.Drawing.Size(298, 28);
            this.errorPane.TabIndex = 9;
            // 
            // settingsPane
            // 
            this.settingsPane.BackColor = System.Drawing.Color.LightSteelBlue;
            this.settingsPane.Location = new System.Drawing.Point(-12, -10);
            this.settingsPane.Name = "settingsPane";
            this.settingsPane.Size = new System.Drawing.Size(294, 222);
            this.settingsPane.TabIndex = 8;
            // 
            // tweetGrid
            // 
            this.tweetGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tweetGrid.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tweetGrid.DataSource = null;
            this.tweetGrid.Enabled = false;
            this.tweetGrid.Location = new System.Drawing.Point(1, 4);
            this.tweetGrid.Margin = new System.Windows.Forms.Padding(4);
            this.tweetGrid.Name = "tweetGrid";
            this.tweetGrid.Size = new System.Drawing.Size(295, 665);
            this.tweetGrid.TabIndex = 7;
            // 
            // TwitterPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.splitContainerButtons);
            this.Controls.Add(this.errorPane);
            this.Controls.Add(this.settingsPane);
            this.Controls.Add(this.tweetGrid);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TwitterPane";
            this.Size = new System.Drawing.Size(300, 735);
            this.splitContainerButtons.Panel1.ResumeLayout(false);
            this.splitContainerButtons.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerButtons)).EndInit();
            this.splitContainerButtons.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonAddinSettings;
        private System.Windows.Forms.Button buttonMain;
        private TweetGrid tweetGrid;
        private SettingsPane settingsPane;
        private ErrorPane errorPane;
        private System.Windows.Forms.SplitContainer splitContainerButtons;
    }
}
