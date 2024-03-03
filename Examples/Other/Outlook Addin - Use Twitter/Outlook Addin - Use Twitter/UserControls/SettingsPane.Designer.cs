namespace Sample.Addin
{
    partial class SettingsPane
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsPane));
            this.label7 = new System.Windows.Forms.Label();
            this.numericRefreshInterval = new System.Windows.Forms.NumericUpDown();
            this.buttonTestConnection = new System.Windows.Forms.Button();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.label6 = new System.Windows.Forms.Label();
            this.linkLabelDeveloperApi = new System.Windows.Forms.LinkLabel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.label5 = new System.Windows.Forms.Label();
            this.panelMessage = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelMessage = new System.Windows.Forms.Label();
            this.textBoxAccessSecret = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxAccessToken = new System.Windows.Forms.TextBox();
            this.textBoxAuthSecret = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxAuthKey = new System.Windows.Forms.TextBox();
            this.labelTestConnection = new System.Windows.Forms.Label();
            this.checkBoxEnabled = new System.Windows.Forms.CheckBox();
            this.label10 = new System.Windows.Forms.Label();
            this.linkLabelLinq2Twitter = new System.Windows.Forms.LinkLabel();
            this.label8 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numericRefreshInterval)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.panelMessage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(25, 293);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(171, 16);
            this.label7.TabIndex = 40;
            this.label7.Text = "Refresh Interval in Seconds";
            // 
            // numericRefreshInterval
            // 
            this.numericRefreshInterval.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.numericRefreshInterval.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.numericRefreshInterval.Location = new System.Drawing.Point(214, 290);
            this.numericRefreshInterval.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numericRefreshInterval.Minimum = new decimal(new int[] {
            90,
            0,
            0,
            0});
            this.numericRefreshInterval.Name = "numericRefreshInterval";
            this.numericRefreshInterval.Size = new System.Drawing.Size(60, 20);
            this.numericRefreshInterval.TabIndex = 39;
            this.numericRefreshInterval.Value = new decimal(new int[] {
            90,
            0,
            0,
            0});
            this.numericRefreshInterval.ValueChanged += new System.EventHandler(this.numericRefreshInterval_ValueChanged);
            // 
            // buttonTestConnection
            // 
            this.buttonTestConnection.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonTestConnection.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonTestConnection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonTestConnection.Image = ((System.Drawing.Image)(resources.GetObject("buttonTestConnection.Image")));
            this.buttonTestConnection.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonTestConnection.Location = new System.Drawing.Point(141, 210);
            this.buttonTestConnection.Margin = new System.Windows.Forms.Padding(4);
            this.buttonTestConnection.Name = "buttonTestConnection";
            this.buttonTestConnection.Size = new System.Drawing.Size(138, 32);
            this.buttonTestConnection.TabIndex = 28;
            this.buttonTestConnection.Text = "      Test connection";
            this.buttonTestConnection.UseVisualStyleBackColor = true;
            this.buttonTestConnection.Click += new System.EventHandler(this.buttonTestConnection_Click);
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
            this.pictureBox3.Location = new System.Drawing.Point(24, 332);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(16, 16);
            this.pictureBox3.TabIndex = 37;
            this.pictureBox3.TabStop = false;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(22, 356);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(118, 16);
            this.label6.TabIndex = 36;
            this.label6.Text = "you have to create";
            // 
            // linkLabelDeveloperApi
            // 
            this.linkLabelDeveloperApi.AutoSize = true;
            this.linkLabelDeveloperApi.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelDeveloperApi.Location = new System.Drawing.Point(46, 331);
            this.linkLabelDeveloperApi.Name = "linkLabelDeveloperApi";
            this.linkLabelDeveloperApi.Size = new System.Drawing.Size(135, 16);
            this.linkLabelDeveloperApi.TabIndex = 35;
            this.linkLabelDeveloperApi.TabStop = true;
            this.linkLabelDeveloperApi.Tag = "https://dev.twitter.com";
            this.linkLabelDeveloperApi.Text = "https://dev.twitter.com";
            this.linkLabelDeveloperApi.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.pictureBox2);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Location = new System.Drawing.Point(24, 114);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(255, 23);
            this.panel2.TabIndex = 34;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(0, 0);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(16, 16);
            this.pictureBox2.TabIndex = 6;
            this.pictureBox2.TabStop = false;
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.DarkBlue;
            this.label5.Location = new System.Drawing.Point(22, 1);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(104, 16);
            this.label5.TabIndex = 5;
            this.label5.Text = "Access Settings";
            // 
            // panelMessage
            // 
            this.panelMessage.Controls.Add(this.pictureBox1);
            this.panelMessage.Controls.Add(this.labelMessage);
            this.panelMessage.Location = new System.Drawing.Point(25, 20);
            this.panelMessage.Name = "panelMessage";
            this.panelMessage.Size = new System.Drawing.Size(254, 23);
            this.panelMessage.TabIndex = 33;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(16, 16);
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // labelMessage
            // 
            this.labelMessage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelMessage.AutoSize = true;
            this.labelMessage.BackColor = System.Drawing.Color.Transparent;
            this.labelMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMessage.ForeColor = System.Drawing.Color.DarkBlue;
            this.labelMessage.Location = new System.Drawing.Point(22, 1);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(187, 16);
            this.labelMessage.TabIndex = 5;
            this.labelMessage.Text = "Authentication Settings (oAuth)";
            // 
            // textBoxAccessSecret
            // 
            this.textBoxAccessSecret.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxAccessSecret.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxAccessSecret.Location = new System.Drawing.Point(141, 172);
            this.textBoxAccessSecret.Name = "textBoxAccessSecret";
            this.textBoxAccessSecret.PasswordChar = '*';
            this.textBoxAccessSecret.Size = new System.Drawing.Size(138, 20);
            this.textBoxAccessSecret.TabIndex = 32;
            this.textBoxAccessSecret.TextChanged += new System.EventHandler(this.textBoxAccessSecret_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 172);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(76, 13);
            this.label3.TabIndex = 31;
            this.label3.Text = "Access Secret";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 146);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 13);
            this.label4.TabIndex = 30;
            this.label4.Text = "Access Token";
            // 
            // textBoxAccessToken
            // 
            this.textBoxAccessToken.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxAccessToken.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxAccessToken.Location = new System.Drawing.Point(141, 144);
            this.textBoxAccessToken.Name = "textBoxAccessToken";
            this.textBoxAccessToken.PasswordChar = '*';
            this.textBoxAccessToken.Size = new System.Drawing.Size(138, 20);
            this.textBoxAccessToken.TabIndex = 29;
            this.textBoxAccessToken.TextChanged += new System.EventHandler(this.textBoxAccessToken_TextChanged);
            // 
            // textBoxAuthSecret
            // 
            this.textBoxAuthSecret.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxAuthSecret.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxAuthSecret.Location = new System.Drawing.Point(141, 80);
            this.textBoxAuthSecret.Name = "textBoxAuthSecret";
            this.textBoxAuthSecret.PasswordChar = '*';
            this.textBoxAuthSecret.Size = new System.Drawing.Size(138, 20);
            this.textBoxAuthSecret.TabIndex = 27;
            this.textBoxAuthSecret.TextChanged += new System.EventHandler(this.textBoxAuthSecret_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 13);
            this.label2.TabIndex = 26;
            this.label2.Text = "Consumer Secret";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 25;
            this.label1.Text = "Consumer Key";
            // 
            // textBoxAuthKey
            // 
            this.textBoxAuthKey.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxAuthKey.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxAuthKey.Location = new System.Drawing.Point(141, 52);
            this.textBoxAuthKey.Name = "textBoxAuthKey";
            this.textBoxAuthKey.PasswordChar = '*';
            this.textBoxAuthKey.Size = new System.Drawing.Size(138, 20);
            this.textBoxAuthKey.TabIndex = 24;
            this.textBoxAuthKey.TextChanged += new System.EventHandler(this.textBoxAuthKey_TextChanged);
            // 
            // labelTestConnection
            // 
            this.labelTestConnection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTestConnection.ForeColor = System.Drawing.Color.Red;
            this.labelTestConnection.Location = new System.Drawing.Point(139, 252);
            this.labelTestConnection.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelTestConnection.Name = "labelTestConnection";
            this.labelTestConnection.Size = new System.Drawing.Size(152, 15);
            this.labelTestConnection.TabIndex = 45;
            this.labelTestConnection.Text = "Authentication failed.";
            this.labelTestConnection.Visible = false;
            // 
            // checkBoxEnabled
            // 
            this.checkBoxEnabled.AutoSize = true;
            this.checkBoxEnabled.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.checkBoxEnabled.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxEnabled.Location = new System.Drawing.Point(24, 217);
            this.checkBoxEnabled.Name = "checkBoxEnabled";
            this.checkBoxEnabled.Size = new System.Drawing.Size(76, 20);
            this.checkBoxEnabled.TabIndex = 46;
            this.checkBoxEnabled.Text = "Enabled";
            this.checkBoxEnabled.UseVisualStyleBackColor = true;
            this.checkBoxEnabled.CheckedChanged += new System.EventHandler(this.checkBoxEnabled_CheckedChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(20, 434);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(254, 18);
            this.label10.TabIndex = 47;
            this.label10.Text = "This example Addin use Linq2Twitter. ";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // linkLabelLinq2Twitter
            // 
            this.linkLabelLinq2Twitter.AutoSize = true;
            this.linkLabelLinq2Twitter.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelLinq2Twitter.Location = new System.Drawing.Point(22, 456);
            this.linkLabelLinq2Twitter.Name = "linkLabelLinq2Twitter";
            this.linkLabelLinq2Twitter.Size = new System.Drawing.Size(215, 18);
            this.linkLabelLinq2Twitter.TabIndex = 48;
            this.linkLabelLinq2Twitter.TabStop = true;
            this.linkLabelLinq2Twitter.Tag = "http://linqtotwitter.codeplex.com";
            this.linkLabelLinq2Twitter.Text = "http://linqtotwitter.codeplex.com";
            this.linkLabelLinq2Twitter.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(22, 377);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(189, 16);
            this.label8.TabIndex = 49;
            this.label8.Text = "an application for the twitter api";
            // 
            // SettingsPane
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.linkLabelLinq2Twitter);
            this.Controls.Add(this.checkBoxEnabled);
            this.Controls.Add(this.labelTestConnection);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.numericRefreshInterval);
            this.Controls.Add(this.buttonTestConnection);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.linkLabelDeveloperApi);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panelMessage);
            this.Controls.Add(this.textBoxAccessSecret);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxAccessToken);
            this.Controls.Add(this.textBoxAuthSecret);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxAuthKey);
            this.Name = "SettingsPane";
            this.Size = new System.Drawing.Size(300, 573);
            ((System.ComponentModel.ISupportInitialize)(this.numericRefreshInterval)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.panelMessage.ResumeLayout(false);
            this.panelMessage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown numericRefreshInterval;
        private System.Windows.Forms.Button buttonTestConnection;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.LinkLabel linkLabelDeveloperApi;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel panelMessage;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.TextBox textBoxAccessSecret;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxAccessToken;
        private System.Windows.Forms.TextBox textBoxAuthSecret;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxAuthKey;
        private System.Windows.Forms.Label labelTestConnection;
        private System.Windows.Forms.CheckBox checkBoxEnabled;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.LinkLabel linkLabelLinq2Twitter;
        private System.Windows.Forms.Label label8;
    }
}
