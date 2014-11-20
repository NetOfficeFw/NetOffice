namespace NetOffice.DeveloperToolbox.Controls.Custom
{
    partial class InviteControl
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InviteControl));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelHint = new System.Windows.Forms.Label();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.panelLeft = new System.Windows.Forms.Panel();
            this.linkLabelMail = new System.Windows.Forms.LinkLabel();
            this.panelRight = new System.Windows.Forms.Panel();
            this.linkLabelFolder = new System.Windows.Forms.LinkLabel();
            this.advRichTextBox2 = new NetOffice.DeveloperToolbox.Controls.Text.AdvRichTextBox();
            this.advRichTextBox1 = new NetOffice.DeveloperToolbox.Controls.Text.AdvRichTextBox();
            this.buttonClose = new NetOffice.DeveloperToolbox.Controls.Buttons.RoundedButton();
            this.controlBackColorAnimator1 = new NetOffice.DeveloperToolbox.Utils.Animation.Colors.ControlBackColorAnimator(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panelBottom.SuspendLayout();
            this.panelLeft.SuspendLayout();
            this.panelRight.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.controlBackColorAnimator1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(1076, 542);
            this.pictureBox1.TabIndex = 105;
            this.pictureBox1.TabStop = false;
            // 
            // labelHint
            // 
            this.labelHint.AutoSize = true;
            this.labelHint.BackColor = System.Drawing.Color.White;
            this.labelHint.Font = new System.Drawing.Font("Segoe UI", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelHint.ForeColor = System.Drawing.Color.Blue;
            this.labelHint.Location = new System.Drawing.Point(55, 48);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(475, 86);
            this.labelHint.TabIndex = 106;
            this.labelHint.Text = "WE WANT YOU";
            // 
            // panelBottom
            // 
            this.panelBottom.BackColor = System.Drawing.Color.White;
            this.panelBottom.Controls.Add(this.buttonClose);
            this.panelBottom.Controls.Add(this.labelHint);
            this.panelBottom.Location = new System.Drawing.Point(265, 403);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(565, 138);
            this.panelBottom.TabIndex = 109;
            // 
            // panelLeft
            // 
            this.panelLeft.BackColor = System.Drawing.Color.White;
            this.panelLeft.Controls.Add(this.linkLabelMail);
            this.panelLeft.Controls.Add(this.advRichTextBox1);
            this.panelLeft.Location = new System.Drawing.Point(35, 26);
            this.panelLeft.Name = "panelLeft";
            this.panelLeft.Size = new System.Drawing.Size(213, 166);
            this.panelLeft.TabIndex = 110;
            // 
            // linkLabelMail
            // 
            this.linkLabelMail.AutoSize = true;
            this.linkLabelMail.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelMail.Location = new System.Drawing.Point(6, 88);
            this.linkLabelMail.Name = "linkLabelMail";
            this.linkLabelMail.Size = new System.Drawing.Size(194, 17);
            this.linkLabelMail.TabIndex = 109;
            this.linkLabelMail.TabStop = true;
            this.linkLabelMail.Text = "mailto:public.sebastian@web.de";
            this.linkLabelMail.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMail_LinkClicked);
            // 
            // panelRight
            // 
            this.panelRight.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.panelRight.BackColor = System.Drawing.Color.White;
            this.panelRight.Controls.Add(this.linkLabelFolder);
            this.panelRight.Controls.Add(this.advRichTextBox2);
            this.panelRight.Location = new System.Drawing.Point(831, 26);
            this.panelRight.Name = "panelRight";
            this.panelRight.Size = new System.Drawing.Size(213, 166);
            this.panelRight.TabIndex = 111;
            // 
            // linkLabelFolder
            // 
            this.linkLabelFolder.AutoSize = true;
            this.linkLabelFolder.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelFolder.Location = new System.Drawing.Point(6, 88);
            this.linkLabelFolder.Name = "linkLabelFolder";
            this.linkLabelFolder.Size = new System.Drawing.Size(158, 17);
            this.linkLabelFolder.TabIndex = 109;
            this.linkLabelFolder.TabStop = true;
            this.linkLabelFolder.Text = "Open Language Directory";
            this.linkLabelFolder.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelFolder_LinkClicked);
            // 
            // advRichTextBox2
            // 
            this.advRichTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.advRichTextBox2.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.advRichTextBox2.ForeColor = System.Drawing.Color.DimGray;
            this.advRichTextBox2.Location = new System.Drawing.Point(7, 15);
            this.advRichTextBox2.Name = "advRichTextBox2";
            this.advRichTextBox2.SelectionAlignment = NetOffice.DeveloperToolbox.Controls.Text.TextAlign.Justify;
            this.advRichTextBox2.Size = new System.Drawing.Size(200, 70);
            this.advRichTextBox2.TabIndex = 108;
            this.advRichTextBox2.Text = "Create a new language and send the language file ({LCID}.lng) to the project.";
            // 
            // advRichTextBox1
            // 
            this.advRichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.advRichTextBox1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.advRichTextBox1.ForeColor = System.Drawing.Color.DimGray;
            this.advRichTextBox1.Location = new System.Drawing.Point(7, 15);
            this.advRichTextBox1.Name = "advRichTextBox1";
            this.advRichTextBox1.SelectionAlignment = NetOffice.DeveloperToolbox.Controls.Text.TextAlign.Justify;
            this.advRichTextBox1.Size = new System.Drawing.Size(200, 70);
            this.advRichTextBox1.TabIndex = 108;
            this.advRichTextBox1.Text = "Join NetOffice and create a new language package. If you need help, don\'t be afra" +
                "id to ask.";
            // 
            // buttonClose
            // 
            this.buttonClose.BackColor = System.Drawing.Color.White;
            this.buttonClose.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonClose.ForeColor = System.Drawing.Color.Blue;
            this.buttonClose.Image = ((System.Drawing.Image)(resources.GetObject("buttonClose.Image")));
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.Location = new System.Drawing.Point(54, 1);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(476, 35);
            this.buttonClose.TabIndex = 107;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = false;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // controlBackColorAnimator1
            // 
            this.controlBackColorAnimator1.Control = this.buttonClose;
            this.controlBackColorAnimator1.EndColor = System.Drawing.Color.CadetBlue;
            this.controlBackColorAnimator1.Intervall = 30;
            this.controlBackColorAnimator1.LoopMode = NetOffice.DeveloperToolbox.Utils.Animation.LoopMode.Bidirectional;
            this.controlBackColorAnimator1.StartColor = System.Drawing.Color.White;
            // 
            // InviteControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.panelRight);
            this.Controls.Add(this.panelLeft);
            this.Controls.Add(this.panelBottom);
            this.Controls.Add(this.pictureBox1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "InviteControl";
            this.Size = new System.Drawing.Size(1076, 542);
            this.Resize += new System.EventHandler(this.InviteControl_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panelBottom.ResumeLayout(false);
            this.panelBottom.PerformLayout();
            this.panelLeft.ResumeLayout(false);
            this.panelLeft.PerformLayout();
            this.panelRight.ResumeLayout(false);
            this.panelRight.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.controlBackColorAnimator1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label labelHint;
        private Buttons.RoundedButton buttonClose;
        private Text.AdvRichTextBox advRichTextBox1;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Panel panelLeft;
        private System.Windows.Forms.LinkLabel linkLabelMail;
        private System.Windows.Forms.Panel panelRight;
        private System.Windows.Forms.LinkLabel linkLabelFolder;
        private Text.AdvRichTextBox advRichTextBox2;
        private Utils.Animation.Colors.ControlBackColorAnimator controlBackColorAnimator1;
    }
}
