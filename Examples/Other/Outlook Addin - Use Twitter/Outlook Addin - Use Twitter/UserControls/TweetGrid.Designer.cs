namespace Sample.Addin
{
    partial class TweetGrid
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TweetGrid));
            this.panelSendTweet = new System.Windows.Forms.Panel();
            this.buttonSendTweet = new System.Windows.Forms.Button();
            this.textBoxTweetContent = new System.Windows.Forms.RichTextBox();
            this.panelTweets = new System.Windows.Forms.Panel();
            this.panelTweetPanels = new System.Windows.Forms.Panel();
            this.vScrollBar = new System.Windows.Forms.VScrollBar();
            this.panelError = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.labelErrorMessage = new System.Windows.Forms.Label();
            this.panelSendTweet.SuspendLayout();
            this.panelTweets.SuspendLayout();
            this.panelError.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelSendTweet
            // 
            this.panelSendTweet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelSendTweet.BackColor = System.Drawing.Color.LightSteelBlue;
            this.panelSendTweet.Controls.Add(this.panelError);
            this.panelSendTweet.Controls.Add(this.buttonSendTweet);
            this.panelSendTweet.Controls.Add(this.textBoxTweetContent);
            this.panelSendTweet.Location = new System.Drawing.Point(3, 377);
            this.panelSendTweet.Name = "panelSendTweet";
            this.panelSendTweet.Size = new System.Drawing.Size(295, 120);
            this.panelSendTweet.TabIndex = 0;
            // 
            // buttonSendTweet
            // 
            this.buttonSendTweet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSendTweet.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonSendTweet.Image = ((System.Drawing.Image)(resources.GetObject("buttonSendTweet.Image")));
            this.buttonSendTweet.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonSendTweet.Location = new System.Drawing.Point(190, 81);
            this.buttonSendTweet.Margin = new System.Windows.Forms.Padding(4);
            this.buttonSendTweet.Name = "buttonSendTweet";
            this.buttonSendTweet.Size = new System.Drawing.Size(103, 37);
            this.buttonSendTweet.TabIndex = 29;
            this.buttonSendTweet.Text = "Tweet";
            this.buttonSendTweet.UseVisualStyleBackColor = true;
            this.buttonSendTweet.Click += new System.EventHandler(this.buttonSendTweet_Click);
            // 
            // textBoxTweetContent
            // 
            this.textBoxTweetContent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxTweetContent.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxTweetContent.Location = new System.Drawing.Point(0, 0);
            this.textBoxTweetContent.MaxLength = 180;
            this.textBoxTweetContent.Name = "textBoxTweetContent";
            this.textBoxTweetContent.Size = new System.Drawing.Size(293, 79);
            this.textBoxTweetContent.TabIndex = 0;
            this.textBoxTweetContent.Text = "";
            this.textBoxTweetContent.TextChanged += new System.EventHandler(this.textBoxTweetContent_TextChanged);
            // 
            // panelTweets
            // 
            this.panelTweets.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelTweets.Controls.Add(this.panelTweetPanels);
            this.panelTweets.Controls.Add(this.vScrollBar);
            this.panelTweets.Location = new System.Drawing.Point(3, 3);
            this.panelTweets.Name = "panelTweets";
            this.panelTweets.Size = new System.Drawing.Size(295, 374);
            this.panelTweets.TabIndex = 1;
            // 
            // panelTweetPanels
            // 
            this.panelTweetPanels.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelTweetPanels.BackColor = System.Drawing.Color.LightSteelBlue;
            this.panelTweetPanels.Location = new System.Drawing.Point(3, 3);
            this.panelTweetPanels.Name = "panelTweetPanels";
            this.panelTweetPanels.Size = new System.Drawing.Size(268, 367);
            this.panelTweetPanels.TabIndex = 1;
            // 
            // vScrollBar
            // 
            this.vScrollBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.vScrollBar.Location = new System.Drawing.Point(274, 0);
            this.vScrollBar.Name = "vScrollBar";
            this.vScrollBar.Size = new System.Drawing.Size(19, 371);
            this.vScrollBar.TabIndex = 0;
            this.vScrollBar.Scroll += new System.Windows.Forms.ScrollEventHandler(this.vScrollBar_Scroll);
            // 
            // panelError
            // 
            this.panelError.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelError.Controls.Add(this.labelErrorMessage);
            this.panelError.Controls.Add(this.label1);
            this.panelError.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.panelError.Location = new System.Drawing.Point(2, 81);
            this.panelError.Name = "panelError";
            this.panelError.Size = new System.Drawing.Size(185, 37);
            this.panelError.TabIndex = 30;
            this.panelError.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "An error ocurred.";
            // 
            // labelErrorMessage
            // 
            this.labelErrorMessage.AutoSize = true;
            this.labelErrorMessage.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.labelErrorMessage.Location = new System.Drawing.Point(4, 19);
            this.labelErrorMessage.Name = "labelErrorMessage";
            this.labelErrorMessage.Size = new System.Drawing.Size(84, 13);
            this.labelErrorMessage.TabIndex = 1;
            this.labelErrorMessage.Text = "<ErrorMessage>";
            // 
            // TweetGrid
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.panelTweets);
            this.Controls.Add(this.panelSendTweet);
            this.Name = "TweetGrid";
            this.Size = new System.Drawing.Size(300, 500);
            this.panelSendTweet.ResumeLayout(false);
            this.panelTweets.ResumeLayout(false);
            this.panelError.ResumeLayout(false);
            this.panelError.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelSendTweet;
        private System.Windows.Forms.RichTextBox textBoxTweetContent;
        private System.Windows.Forms.Button buttonSendTweet;
        private System.Windows.Forms.Panel panelTweets;
        private System.Windows.Forms.VScrollBar vScrollBar;
        private System.Windows.Forms.Panel panelTweetPanels;
        private System.Windows.Forms.Panel panelError;
        private System.Windows.Forms.Label labelErrorMessage;
        private System.Windows.Forms.Label label1;

    }
}
