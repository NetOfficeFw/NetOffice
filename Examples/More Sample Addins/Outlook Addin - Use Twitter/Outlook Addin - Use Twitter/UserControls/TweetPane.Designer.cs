namespace Sample.Addin
{
    partial class TweetPane
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
            this.richTextBoxMessage = new System.Windows.Forms.RichTextBox();
            this.pictureBoxImage = new System.Windows.Forms.PictureBox();
            this.linkLabelReply = new System.Windows.Forms.LinkLabel();
            this.linkLabelRetweet = new System.Windows.Forms.LinkLabel();
            this.linklabelFavorite = new System.Windows.Forms.LinkLabel();
            this.labelUserName = new System.Windows.Forms.Label();
            this.labelCreated = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // richTextBoxMessage
            // 
            this.richTextBoxMessage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.richTextBoxMessage.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBoxMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBoxMessage.Location = new System.Drawing.Point(49, 36);
            this.richTextBoxMessage.Margin = new System.Windows.Forms.Padding(4);
            this.richTextBoxMessage.Name = "richTextBoxMessage";
            this.richTextBoxMessage.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
            this.richTextBoxMessage.Size = new System.Drawing.Size(210, 67);
            this.richTextBoxMessage.TabIndex = 20;
            this.richTextBoxMessage.Text = "";
            this.richTextBoxMessage.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.richTextBoxMessage_LinkClicked);
            this.richTextBoxMessage.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBoxMessage_KeyPress);
            // 
            // pictureBoxImage
            // 
            this.pictureBoxImage.Location = new System.Drawing.Point(-1, 36);
            this.pictureBoxImage.Margin = new System.Windows.Forms.Padding(4);
            this.pictureBoxImage.Name = "pictureBoxImage";
            this.pictureBoxImage.Size = new System.Drawing.Size(48, 48);
            this.pictureBoxImage.TabIndex = 21;
            this.pictureBoxImage.TabStop = false;
            // 
            // linkLabelReply
            // 
            this.linkLabelReply.AutoSize = true;
            this.linkLabelReply.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelReply.Location = new System.Drawing.Point(19, 111);
            this.linkLabelReply.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.linkLabelReply.Name = "linkLabelReply";
            this.linkLabelReply.Size = new System.Drawing.Size(49, 20);
            this.linkLabelReply.TabIndex = 22;
            this.linkLabelReply.TabStop = true;
            this.linkLabelReply.Text = "Reply";
            this.linkLabelReply.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelReply_LinkClicked);
            // 
            // linkLabelRetweet
            // 
            this.linkLabelRetweet.AutoSize = true;
            this.linkLabelRetweet.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelRetweet.Location = new System.Drawing.Point(89, 110);
            this.linkLabelRetweet.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.linkLabelRetweet.Name = "linkLabelRetweet";
            this.linkLabelRetweet.Size = new System.Drawing.Size(69, 20);
            this.linkLabelRetweet.TabIndex = 24;
            this.linkLabelRetweet.TabStop = true;
            this.linkLabelRetweet.Text = "Retweet";
            this.linkLabelRetweet.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelRetweet_LinkClicked);
            // 
            // linklabelFavorite
            // 
            this.linklabelFavorite.AutoSize = true;
            this.linklabelFavorite.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linklabelFavorite.Location = new System.Drawing.Point(183, 110);
            this.linklabelFavorite.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.linklabelFavorite.Name = "linklabelFavorite";
            this.linklabelFavorite.Size = new System.Drawing.Size(66, 20);
            this.linklabelFavorite.TabIndex = 25;
            this.linklabelFavorite.TabStop = true;
            this.linklabelFavorite.Text = "Favorite";
            this.linklabelFavorite.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linklabelFavorite_LinkClicked);
            // 
            // labelUserName
            // 
            this.labelUserName.AutoSize = true;
            this.labelUserName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelUserName.Location = new System.Drawing.Point(4, 14);
            this.labelUserName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelUserName.Name = "labelUserName";
            this.labelUserName.Size = new System.Drawing.Size(102, 16);
            this.labelUserName.TabIndex = 26;
            this.labelUserName.Text = "DisplayName";
            // 
            // labelCreated
            // 
            this.labelCreated.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelCreated.AutoSize = true;
            this.labelCreated.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCreated.Location = new System.Drawing.Point(153, 16);
            this.labelCreated.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelCreated.Name = "labelCreated";
            this.labelCreated.Size = new System.Drawing.Size(106, 13);
            this.labelCreated.TabIndex = 27;
            this.labelCreated.Text = "10.10.2010 10:10:10";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(0, 84);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(24, 19);
            this.pictureBox1.TabIndex = 28;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Location = new System.Drawing.Point(24, 84);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(23, 19);
            this.pictureBox2.TabIndex = 29;
            this.pictureBox2.TabStop = false;
            // 
            // TweetPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.labelCreated);
            this.Controls.Add(this.labelUserName);
            this.Controls.Add(this.linklabelFavorite);
            this.Controls.Add(this.linkLabelRetweet);
            this.Controls.Add(this.linkLabelReply);
            this.Controls.Add(this.pictureBoxImage);
            this.Controls.Add(this.richTextBoxMessage);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TweetPane";
            this.Size = new System.Drawing.Size(260, 140);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox richTextBoxMessage;
        private System.Windows.Forms.PictureBox pictureBoxImage;
        private System.Windows.Forms.LinkLabel linkLabelReply;
        private System.Windows.Forms.LinkLabel linkLabelRetweet;
        private System.Windows.Forms.LinkLabel linklabelFavorite;
        private System.Windows.Forms.Label labelUserName;
        private System.Windows.Forms.Label labelCreated;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
    }
}
