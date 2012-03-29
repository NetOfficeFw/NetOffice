namespace NetOffice.ProjectWizard
{
    partial class ErrorDialog
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorDialog));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelErrorCaption = new System.Windows.Forms.Label();
            this.okayButton = new System.Windows.Forms.Button();
            this.copyButton = new System.Windows.Forms.Button();
            this.textBoxErrorLog = new System.Windows.Forms.TextBox();
            this.labelErrorLog = new System.Windows.Forms.Label();
            this.labelHint = new System.Windows.Forms.Label();
            this.linkLabelNetOfficeDiscussion = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(40, 25);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(34, 33);
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            // 
            // labelErrorCaption
            // 
            this.labelErrorCaption.AutoSize = true;
            this.labelErrorCaption.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorCaption.Location = new System.Drawing.Point(109, 25);
            this.labelErrorCaption.Name = "labelErrorCaption";
            this.labelErrorCaption.Size = new System.Drawing.Size(268, 15);
            this.labelErrorCaption.TabIndex = 9;
            this.labelErrorCaption.Text = "Leider ist ein unerwarteter Fehler aufgetreten.";
            // 
            // okayButton
            // 
            this.okayButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okayButton.Location = new System.Drawing.Point(312, 338);
            this.okayButton.Name = "okayButton";
            this.okayButton.Size = new System.Drawing.Size(95, 25);
            this.okayButton.TabIndex = 10;
            this.okayButton.Text = "OK";
            this.okayButton.UseVisualStyleBackColor = true;
            this.okayButton.Click += new System.EventHandler(this.okayButton_Click);
            // 
            // copyButton
            // 
            this.copyButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.copyButton.Location = new System.Drawing.Point(24, 338);
            this.copyButton.Name = "copyButton";
            this.copyButton.Size = new System.Drawing.Size(244, 25);
            this.copyButton.TabIndex = 11;
            this.copyButton.Text = "Fehlerbericht in die Zwischenablage kopieren";
            this.copyButton.UseVisualStyleBackColor = true;
            this.copyButton.Click += new System.EventHandler(this.copyButton_Click);
            // 
            // textBoxErrorLog
            // 
            this.textBoxErrorLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxErrorLog.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.textBoxErrorLog.Location = new System.Drawing.Point(24, 118);
            this.textBoxErrorLog.Multiline = true;
            this.textBoxErrorLog.Name = "textBoxErrorLog";
            this.textBoxErrorLog.ReadOnly = true;
            this.textBoxErrorLog.Size = new System.Drawing.Size(384, 204);
            this.textBoxErrorLog.TabIndex = 12;
            // 
            // labelErrorLog
            // 
            this.labelErrorLog.AutoSize = true;
            this.labelErrorLog.Location = new System.Drawing.Point(21, 102);
            this.labelErrorLog.Name = "labelErrorLog";
            this.labelErrorLog.Size = new System.Drawing.Size(76, 13);
            this.labelErrorLog.TabIndex = 13;
            this.labelErrorLog.Text = "Fehlerprotokoll";
            // 
            // labelHint
            // 
            this.labelHint.Location = new System.Drawing.Point(109, 52);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(295, 30);
            this.labelHint.TabIndex = 14;
            this.labelHint.Text = "Wenn Sie möchten können Sie das NetOffice Forum nutzen um weitere Hilfe zu erhalt" +
                "en.";
            // 
            // linkLabelNetOfficeDiscussion
            // 
            this.linkLabelNetOfficeDiscussion.AutoSize = true;
            this.linkLabelNetOfficeDiscussion.Location = new System.Drawing.Point(249, 65);
            this.linkLabelNetOfficeDiscussion.Name = "linkLabelNetOfficeDiscussion";
            this.linkLabelNetOfficeDiscussion.Size = new System.Drawing.Size(137, 13);
            this.linkLabelNetOfficeDiscussion.TabIndex = 15;
            this.linkLabelNetOfficeDiscussion.TabStop = true;
            this.linkLabelNetOfficeDiscussion.Text = "NetOffice Discussion Board";
            this.linkLabelNetOfficeDiscussion.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelNetOfficeDiscussion_LinkClicked);
            // 
            // ErrorDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(435, 378);
            this.Controls.Add(this.linkLabelNetOfficeDiscussion);
            this.Controls.Add(this.labelHint);
            this.Controls.Add(this.labelErrorLog);
            this.Controls.Add(this.textBoxErrorLog);
            this.Controls.Add(this.copyButton);
            this.Controls.Add(this.okayButton);
            this.Controls.Add(this.labelErrorCaption);
            this.Controls.Add(this.pictureBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ErrorDialog";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Fehler";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label labelErrorCaption;
        private System.Windows.Forms.Button okayButton;
        private System.Windows.Forms.Button copyButton;
        private System.Windows.Forms.TextBox textBoxErrorLog;
        private System.Windows.Forms.Label labelErrorLog;
        private System.Windows.Forms.Label labelHint;
        private System.Windows.Forms.LinkLabel linkLabelNetOfficeDiscussion;

    }
}
