namespace NetOffice.OfficeApi.Tools.Dialogs
{
    partial class AboutDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutDialog));
            this.panelHeader = new System.Windows.Forms.Panel();
            this.labelAbout = new System.Windows.Forms.Label();
            this.pictureBoxHeader = new System.Windows.Forms.PictureBox();
            this.labelAssemblyTitleVersion = new System.Windows.Forms.Label();
            this.buttonClose = new System.Windows.Forms.Button();
            this.linkLabelCompany = new System.Windows.Forms.LinkLabel();
            this.labelCopyright = new System.Windows.Forms.Label();
            this.richTextBoxLicence = new System.Windows.Forms.RichTextBox();
            this.panelLicenceContent = new System.Windows.Forms.Panel();
            this.panelLicenceHeader = new System.Windows.Forms.Panel();
            this.labelLicenceHeader = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).BeginInit();
            this.panelLicenceContent.SuspendLayout();
            this.panelLicenceHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panelHeader
            // 
            this.panelHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelHeader.BackColor = System.Drawing.Color.White;
            this.panelHeader.Controls.Add(this.labelAbout);
            this.panelHeader.Controls.Add(this.pictureBoxHeader);
            this.panelHeader.Controls.Add(this.labelAssemblyTitleVersion);
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(432, 58);
            this.panelHeader.TabIndex = 3;
            // 
            // labelAbout
            // 
            this.labelAbout.AutoSize = true;
            this.labelAbout.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelAbout.ForeColor = System.Drawing.Color.Black;
            this.labelAbout.Location = new System.Drawing.Point(73, 12);
            this.labelAbout.Name = "labelAbout";
            this.labelAbout.Size = new System.Drawing.Size(61, 16);
            this.labelAbout.TabIndex = 1;
            this.labelAbout.Text = "%About";
            // 
            // pictureBoxHeader
            // 
            this.pictureBoxHeader.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxHeader.Image")));
            this.pictureBoxHeader.Location = new System.Drawing.Point(25, 13);
            this.pictureBoxHeader.Name = "pictureBoxHeader";
            this.pictureBoxHeader.Size = new System.Drawing.Size(34, 34);
            this.pictureBoxHeader.TabIndex = 0;
            this.pictureBoxHeader.TabStop = false;
            // 
            // labelAssemblyTitleVersion
            // 
            this.labelAssemblyTitleVersion.AutoSize = true;
            this.labelAssemblyTitleVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelAssemblyTitleVersion.ForeColor = System.Drawing.Color.Gray;
            this.labelAssemblyTitleVersion.Location = new System.Drawing.Point(73, 31);
            this.labelAssemblyTitleVersion.Name = "labelAssemblyTitleVersion";
            this.labelAssemblyTitleVersion.Size = new System.Drawing.Size(239, 16);
            this.labelAssemblyTitleVersion.TabIndex = 6;
            this.labelAssemblyTitleVersion.Text = "%AssemblyTitle && %AssemblyVersion";
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonClose.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.ForeColor = System.Drawing.Color.Blue;
            this.buttonClose.Image = ((System.Drawing.Image)(resources.GetObject("buttonClose.Image")));
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.Location = new System.Drawing.Point(284, 332);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(118, 29);
            this.buttonClose.TabIndex = 5;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // linkLabelCompany
            // 
            this.linkLabelCompany.AutoSize = true;
            this.linkLabelCompany.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelCompany.Location = new System.Drawing.Point(27, 302);
            this.linkLabelCompany.Name = "linkLabelCompany";
            this.linkLabelCompany.Size = new System.Drawing.Size(78, 16);
            this.linkLabelCompany.TabIndex = 7;
            this.linkLabelCompany.TabStop = true;
            this.linkLabelCompany.Text = "%Company";
            this.linkLabelCompany.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelCompany_LinkClicked);
            // 
            // labelCopyright
            // 
            this.labelCopyright.AutoSize = true;
            this.labelCopyright.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCopyright.ForeColor = System.Drawing.Color.Gray;
            this.labelCopyright.Location = new System.Drawing.Point(28, 73);
            this.labelCopyright.Name = "labelCopyright";
            this.labelCopyright.Size = new System.Drawing.Size(77, 16);
            this.labelCopyright.TabIndex = 8;
            this.labelCopyright.Text = "%Copyright";
            // 
            // richTextBoxLicence
            // 
            this.richTextBoxLicence.BackColor = System.Drawing.Color.White;
            this.richTextBoxLicence.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBoxLicence.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBoxLicence.Location = new System.Drawing.Point(0, 0);
            this.richTextBoxLicence.Name = "richTextBoxLicence";
            this.richTextBoxLicence.ReadOnly = true;
            this.richTextBoxLicence.Size = new System.Drawing.Size(376, 143);
            this.richTextBoxLicence.TabIndex = 9;
            this.richTextBoxLicence.Text = "";
            this.richTextBoxLicence.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.richTextBoxLicence_LinkClicked);
            // 
            // panelLicenceContent
            // 
            this.panelLicenceContent.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelLicenceContent.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.panelLicenceContent.Controls.Add(this.richTextBoxLicence);
            this.panelLicenceContent.Location = new System.Drawing.Point(0, 36);
            this.panelLicenceContent.Name = "panelLicenceContent";
            this.panelLicenceContent.Size = new System.Drawing.Size(376, 143);
            this.panelLicenceContent.TabIndex = 10;
            // 
            // panelLicenceHeader
            // 
            this.panelLicenceHeader.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelLicenceHeader.BackColor = System.Drawing.Color.Orange;
            this.panelLicenceHeader.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelLicenceHeader.Controls.Add(this.labelLicenceHeader);
            this.panelLicenceHeader.Controls.Add(this.pictureBox1);
            this.panelLicenceHeader.Controls.Add(this.panelLicenceContent);
            this.panelLicenceHeader.Location = new System.Drawing.Point(25, 105);
            this.panelLicenceHeader.Name = "panelLicenceHeader";
            this.panelLicenceHeader.Size = new System.Drawing.Size(378, 181);
            this.panelLicenceHeader.TabIndex = 11;
            // 
            // labelLicenceHeader
            // 
            this.labelLicenceHeader.AutoSize = true;
            this.labelLicenceHeader.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelLicenceHeader.Location = new System.Drawing.Point(39, 10);
            this.labelLicenceHeader.Name = "labelLicenceHeader";
            this.labelLicenceHeader.Size = new System.Drawing.Size(131, 17);
            this.labelLicenceHeader.TabIndex = 12;
            this.labelLicenceHeader.Text = "Licence Information";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(6, 6);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(26, 26);
            this.pictureBox1.TabIndex = 12;
            this.pictureBox1.TabStop = false;
            // 
            // AboutDialog
            // 
            this.AcceptButton = this.buttonClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonClose;
            this.ClientSize = new System.Drawing.Size(432, 380);
            this.Controls.Add(this.panelLicenceHeader);
            this.Controls.Add(this.labelCopyright);
            this.Controls.Add(this.linkLabelCompany);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.panelHeader);
            this.MinimumSize = new System.Drawing.Size(432, 375);
            this.Name = "AboutDialog";
            this.Text = "About";
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).EndInit();
            this.panelLicenceContent.ResumeLayout(false);
            this.panelLicenceHeader.ResumeLayout(false);
            this.panelLicenceHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label labelAbout;
        private System.Windows.Forms.PictureBox pictureBoxHeader;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Label labelAssemblyTitleVersion;
        private System.Windows.Forms.LinkLabel linkLabelCompany;
        private System.Windows.Forms.Label labelCopyright;
        private System.Windows.Forms.RichTextBox richTextBoxLicence;
        private System.Windows.Forms.Panel panelLicenceContent;
        private System.Windows.Forms.Panel panelLicenceHeader;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label labelLicenceHeader;
    }
}