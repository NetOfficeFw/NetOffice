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
            this.settingsPane = new Sample.Addin.SettingsPane();
            this.tweetGrid = new Sample.Addin.TweetGrid();
            this.SuspendLayout();
            // 
            // buttonAddinSettings
            // 
            this.buttonAddinSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonAddinSettings.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonAddinSettings.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonAddinSettings.Image = ((System.Drawing.Image)(resources.GetObject("buttonAddinSettings.Image")));
            this.buttonAddinSettings.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonAddinSettings.Location = new System.Drawing.Point(149, 694);
            this.buttonAddinSettings.Margin = new System.Windows.Forms.Padding(4);
            this.buttonAddinSettings.Name = "buttonAddinSettings";
            this.buttonAddinSettings.Size = new System.Drawing.Size(150, 41);
            this.buttonAddinSettings.TabIndex = 5;
            this.buttonAddinSettings.Text = "   Settings";
            this.buttonAddinSettings.UseVisualStyleBackColor = true;
            this.buttonAddinSettings.Click += new System.EventHandler(this.button_Click);
            // 
            // buttonMain
            // 
            this.buttonMain.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonMain.BackColor = System.Drawing.Color.Goldenrod;
            this.buttonMain.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonMain.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonMain.Image = ((System.Drawing.Image)(resources.GetObject("buttonMain.Image")));
            this.buttonMain.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonMain.Location = new System.Drawing.Point(0, 694);
            this.buttonMain.Margin = new System.Windows.Forms.Padding(4);
            this.buttonMain.Name = "buttonMain";
            this.buttonMain.Size = new System.Drawing.Size(150, 41);
            this.buttonMain.TabIndex = 6;
            this.buttonMain.Text = "Main Page";
            this.buttonMain.UseVisualStyleBackColor = false;
            this.buttonMain.Click += new System.EventHandler(this.button_Click);
            // 
            // settingsPane
            // 
            this.settingsPane.BackColor = System.Drawing.Color.LightSteelBlue;
            this.settingsPane.Location = new System.Drawing.Point(3, 186);
            this.settingsPane.Name = "settingsPane";
            this.settingsPane.Size = new System.Drawing.Size(294, 500);
            this.settingsPane.TabIndex = 8;
            // 
            // tweetGrid
            // 
            this.tweetGrid.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tweetGrid.DataSource = null;
            this.tweetGrid.Location = new System.Drawing.Point(1, 4);
            this.tweetGrid.Margin = new System.Windows.Forms.Padding(4);
            this.tweetGrid.Name = "tweetGrid";
            this.tweetGrid.Size = new System.Drawing.Size(295, 175);
            this.tweetGrid.TabIndex = 7;
            // 
            // TwitterPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.settingsPane);
            this.Controls.Add(this.tweetGrid);
            this.Controls.Add(this.buttonMain);
            this.Controls.Add(this.buttonAddinSettings);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TwitterPane";
            this.Size = new System.Drawing.Size(300, 735);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonAddinSettings;
        private System.Windows.Forms.Button buttonMain;
        private TweetGrid tweetGrid;
        private SettingsPane settingsPane;
    }
}
