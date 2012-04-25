namespace NetOffice.DeveloperToolbox.OutlookSecurity
{
    partial class OutlookSecurityControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OutlookSecurityControl));
            this.checkBoxSupressEnabled = new System.Windows.Forms.CheckBox();
            this.labelMessages = new System.Windows.Forms.TextBox();
            this.labelLog = new System.Windows.Forms.Label();
            this.buttonInfo = new System.Windows.Forms.Button();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.labelCaption = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelHint = new System.Windows.Forms.Label();
            this.labelSources = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.linkLabel3 = new System.Windows.Forms.LinkLabel();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.labelDidYouKnow = new System.Windows.Forms.Label();
            this.labelSupressThisDialog = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // checkBoxSupressEnabled
            // 
            this.checkBoxSupressEnabled.AutoSize = true;
            this.checkBoxSupressEnabled.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxSupressEnabled.Location = new System.Drawing.Point(23, 50);
            this.checkBoxSupressEnabled.Name = "checkBoxSupressEnabled";
            this.checkBoxSupressEnabled.Size = new System.Drawing.Size(197, 20);
            this.checkBoxSupressEnabled.TabIndex = 21;
            this.checkBoxSupressEnabled.Text = "Unterdrückung eingeschaltet";
            this.checkBoxSupressEnabled.UseVisualStyleBackColor = true;
            this.checkBoxSupressEnabled.CheckedChanged += new System.EventHandler(this.checkBoxSupressEnabeld_CheckedChanged);
            // 
            // labelMessages
            // 
            this.labelMessages.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelMessages.BackColor = System.Drawing.SystemColors.Control;
            this.labelMessages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelMessages.ForeColor = System.Drawing.Color.White;
            this.labelMessages.Location = new System.Drawing.Point(23, 313);
            this.labelMessages.Multiline = true;
            this.labelMessages.Name = "labelMessages";
            this.labelMessages.ReadOnly = true;
            this.labelMessages.Size = new System.Drawing.Size(758, 96);
            this.labelMessages.TabIndex = 68;
            this.labelMessages.Text = "<Keine Einträge>";
            this.labelMessages.WordWrap = false;
            // 
            // labelLog
            // 
            this.labelLog.AutoSize = true;
            this.labelLog.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelLog.Location = new System.Drawing.Point(20, 297);
            this.labelLog.Name = "labelLog";
            this.labelLog.Size = new System.Drawing.Size(57, 13);
            this.labelLog.TabIndex = 67;
            this.labelLog.Text = "Protokoll";
            // 
            // buttonInfo
            // 
            this.buttonInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonInfo.Image = ((System.Drawing.Image)(resources.GetObject("buttonInfo.Image")));
            this.buttonInfo.Location = new System.Drawing.Point(753, 10);
            this.buttonInfo.Name = "buttonInfo";
            this.buttonInfo.Size = new System.Drawing.Size(28, 28);
            this.buttonInfo.TabIndex = 69;
            this.buttonInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonInfo.UseVisualStyleBackColor = true;
            this.buttonInfo.Click += new System.EventHandler(this.buttonInfo_Click);
            // 
            // pictureBox4
            // 
            this.pictureBox4.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox4.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox4.Image")));
            this.pictureBox4.Location = new System.Drawing.Point(23, 19);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(16, 16);
            this.pictureBox4.TabIndex = 71;
            this.pictureBox4.TabStop = false;
            // 
            // labelCaption
            // 
            this.labelCaption.BackColor = System.Drawing.Color.Khaki;
            this.labelCaption.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelCaption.Location = new System.Drawing.Point(42, 21);
            this.labelCaption.Name = "labelCaption";
            this.labelCaption.Size = new System.Drawing.Size(265, 14);
            this.labelCaption.TabIndex = 70;
            this.labelCaption.Text = "Unterdrücken Sie die Security Abfrage in Outlook.";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(22, 76);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(357, 190);
            this.pictureBox1.TabIndex = 72;
            this.pictureBox1.TabStop = false;
            // 
            // labelHint
            // 
            this.labelHint.Location = new System.Drawing.Point(406, 108);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(285, 27);
            this.labelHint.TabIndex = 73;
            this.labelHint.Text = "Outlook bietet je nach Version verschiedene Wege an diesen Dialog für Ihr Addin n" +
                "icht anzuzeigen.";
            // 
            // labelSources
            // 
            this.labelSources.AutoSize = true;
            this.labelSources.Location = new System.Drawing.Point(406, 144);
            this.labelSources.Name = "labelSources";
            this.labelSources.Size = new System.Drawing.Size(101, 13);
            this.labelSources.TabIndex = 74;
            this.labelSources.Text = "Informationsquellen:";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(407, 174);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(340, 13);
            this.linkLabel1.TabIndex = 75;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "http://msdn.microsoft.com/library/1thd35d7%28en-US,VS.80%29.aspx";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(407, 200);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(353, 13);
            this.linkLabel2.TabIndex = 76;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "http://msdn.microsoft.com/en-US/library/zdc263t0%28v=VS.80%29.aspx";
            // 
            // linkLabel3
            // 
            this.linkLabel3.AutoSize = true;
            this.linkLabel3.Location = new System.Drawing.Point(407, 227);
            this.linkLabel3.Name = "linkLabel3";
            this.linkLabel3.Size = new System.Drawing.Size(365, 13);
            this.linkLabel3.TabIndex = 77;
            this.linkLabel3.TabStop = true;
            this.linkLabel3.Text = "http://msdn.microsoft.com/en-us/library/aa168346%28v=office.11%29.aspx";
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(407, 77);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(16, 16);
            this.pictureBox2.TabIndex = 79;
            this.pictureBox2.TabStop = false;
            // 
            // labelDidYouKnow
            // 
            this.labelDidYouKnow.AutoSize = true;
            this.labelDidYouKnow.BackColor = System.Drawing.SystemColors.Control;
            this.labelDidYouKnow.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelDidYouKnow.Location = new System.Drawing.Point(426, 79);
            this.labelDidYouKnow.Name = "labelDidYouKnow";
            this.labelDidYouKnow.Size = new System.Drawing.Size(117, 13);
            this.labelDidYouKnow.TabIndex = 78;
            this.labelDidYouKnow.Text = "Wussten Sie schon... ?";
            // 
            // labelSupressThisDialog
            // 
            this.labelSupressThisDialog.BackColor = System.Drawing.Color.Transparent;
            this.labelSupressThisDialog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelSupressThisDialog.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSupressThisDialog.ForeColor = System.Drawing.Color.Red;
            this.labelSupressThisDialog.Location = new System.Drawing.Point(140, 80);
            this.labelSupressThisDialog.Name = "labelSupressThisDialog";
            this.labelSupressThisDialog.Size = new System.Drawing.Size(213, 19);
            this.labelSupressThisDialog.TabIndex = 80;
            this.labelSupressThisDialog.Text = "Diesen Dialog unterdrücken";
            // 
            // OutlookSecurityControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelSupressThisDialog);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.labelDidYouKnow);
            this.Controls.Add(this.linkLabel3);
            this.Controls.Add(this.linkLabel2);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.labelSources);
            this.Controls.Add(this.labelHint);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.pictureBox4);
            this.Controls.Add(this.labelCaption);
            this.Controls.Add(this.buttonInfo);
            this.Controls.Add(this.labelMessages);
            this.Controls.Add(this.labelLog);
            this.Controls.Add(this.checkBoxSupressEnabled);
            this.Name = "OutlookSecurityControl";
            this.Size = new System.Drawing.Size(800, 429);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkBoxSupressEnabled;
        private System.Windows.Forms.TextBox labelMessages;
        private System.Windows.Forms.Label labelLog;
        private System.Windows.Forms.Button buttonInfo;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.Label labelCaption;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label labelHint;
        private System.Windows.Forms.Label labelSources;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.LinkLabel linkLabel3;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label labelDidYouKnow;
        private System.Windows.Forms.Label labelSupressThisDialog;
    }
}
