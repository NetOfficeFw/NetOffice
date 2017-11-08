namespace NetOffice.DeveloperToolbox.ToolboxControls.OutlookSecurity
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OutlookSecurityControl));
            this.checkBoxSupressEnabled = new System.Windows.Forms.CheckBox();
            this.labelMessages = new System.Windows.Forms.TextBox();
            this.labelLog = new System.Windows.Forms.Label();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.labelCaption = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelHint = new System.Windows.Forms.Label();
            this.labelSources = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.LinkContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.linkLabel3 = new System.Windows.Forms.LinkLabel();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.labelDidYouKnow = new System.Windows.Forms.Label();
            this.labelSupressThisDialog = new System.Windows.Forms.Label();
            this.linksToolTip = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.LinkContextMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // checkBoxSupressEnabled
            // 
            this.checkBoxSupressEnabled.AutoSize = true;
            this.checkBoxSupressEnabled.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxSupressEnabled.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxSupressEnabled.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxSupressEnabled.Location = new System.Drawing.Point(23, 31);
            this.checkBoxSupressEnabled.Name = "checkBoxSupressEnabled";
            this.checkBoxSupressEnabled.Size = new System.Drawing.Size(135, 20);
            this.checkBoxSupressEnabled.TabIndex = 21;
            this.checkBoxSupressEnabled.Text = "Suppress enabled";
            this.checkBoxSupressEnabled.UseVisualStyleBackColor = true;
            this.checkBoxSupressEnabled.CheckedChanged += new System.EventHandler(this.checkBoxSupressEnabeld_CheckedChanged);
            // 
            // labelMessages
            // 
            this.labelMessages.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelMessages.BackColor = System.Drawing.Color.LightSteelBlue;
            this.labelMessages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelMessages.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMessages.ForeColor = System.Drawing.Color.Black;
            this.labelMessages.Location = new System.Drawing.Point(23, 279);
            this.labelMessages.Multiline = true;
            this.labelMessages.Name = "labelMessages";
            this.labelMessages.ReadOnly = true;
            this.labelMessages.Size = new System.Drawing.Size(882, 192);
            this.labelMessages.TabIndex = 68;
            this.labelMessages.Text = "<No Entries>";
            this.labelMessages.WordWrap = false;
            // 
            // labelLog
            // 
            this.labelLog.AutoSize = true;
            this.labelLog.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelLog.ForeColor = System.Drawing.Color.Black;
            this.labelLog.Location = new System.Drawing.Point(20, 263);
            this.labelLog.Name = "labelLog";
            this.labelLog.Size = new System.Drawing.Size(58, 16);
            this.labelLog.TabIndex = 67;
            this.labelLog.Text = "Protocol";
            // 
            // pictureBox4
            // 
            this.pictureBox4.BackColor = System.Drawing.Color.Orange;
            this.pictureBox4.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox4.Image")));
            this.pictureBox4.Location = new System.Drawing.Point(23, 0);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(20, 16);
            this.pictureBox4.TabIndex = 71;
            this.pictureBox4.TabStop = false;
            // 
            // labelCaption
            // 
            this.labelCaption.BackColor = System.Drawing.Color.Orange;
            this.labelCaption.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCaption.ForeColor = System.Drawing.Color.Black;
            this.labelCaption.Location = new System.Drawing.Point(42, 0);
            this.labelCaption.Name = "labelCaption";
            this.labelCaption.Size = new System.Drawing.Size(337, 16);
            this.labelCaption.TabIndex = 70;
            this.labelCaption.Text = "Suppress the Security Conformation in Outlook.";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(22, 57);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(357, 190);
            this.pictureBox1.TabIndex = 72;
            this.pictureBox1.TabStop = false;
            // 
            // labelHint
            // 
            this.labelHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelHint.ForeColor = System.Drawing.Color.Black;
            this.labelHint.Location = new System.Drawing.Point(406, 89);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(432, 36);
            this.labelHint.TabIndex = 73;
            this.labelHint.Text = "Outlook provides -depending on the version- several ways to suppress this dialog " +
    "for your Addin.";
            // 
            // labelSources
            // 
            this.labelSources.AutoSize = true;
            this.labelSources.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSources.ForeColor = System.Drawing.Color.Black;
            this.labelSources.Location = new System.Drawing.Point(406, 133);
            this.labelSources.Name = "labelSources";
            this.labelSources.Size = new System.Drawing.Size(143, 16);
            this.labelSources.TabIndex = 74;
            this.labelSources.Text = "Sources of information:";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.ContextMenuStrip = this.LinkContextMenu;
            this.linkLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel1.Location = new System.Drawing.Point(407, 155);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(418, 16);
            this.linkLabel1.TabIndex = 75;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Tag = "http://msdn.microsoft.com/library/1thd35d7%28en-US,VS.80%29.aspx";
            this.linkLabel1.Text = "http://msdn.microsoft.com/library/1thd35d7%28en-US,VS.80%29.aspx";
            this.linksToolTip.SetToolTip(this.linkLabel1, "http://msdn.microsoft.com/library/1thd35d7%28en-US,VS.80%29.aspx");
            this.linkLabel1.Click += new System.EventHandler(this.linkLabel_Clicked);
            // 
            // LinkContextMenu
            // 
            this.LinkContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.LinkContextMenu.Name = "LinkContextMenu";
            this.LinkContextMenu.Size = new System.Drawing.Size(177, 26);
            this.LinkContextMenu.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.LinkContextMenu_ItemClicked);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(176, 22);
            this.toolStripMenuItem1.Text = "Copy Link Location";
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.ContextMenuStrip = this.LinkContextMenu;
            this.linkLabel2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel2.Location = new System.Drawing.Point(407, 181);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(431, 16);
            this.linkLabel2.TabIndex = 76;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Tag = "http://msdn.microsoft.com/en-US/library/zdc263t0%28v=VS.80%29.aspx";
            this.linkLabel2.Text = "http://msdn.microsoft.com/en-US/library/zdc263t0%28v=VS.80%29.aspx";
            this.linksToolTip.SetToolTip(this.linkLabel2, "http://msdn.microsoft.com/en-US/library/zdc263t0%28v=VS.80%29.aspx");
            this.linkLabel2.Click += new System.EventHandler(this.linkLabel_Clicked);
            // 
            // linkLabel3
            // 
            this.linkLabel3.AutoSize = true;
            this.linkLabel3.ContextMenuStrip = this.LinkContextMenu;
            this.linkLabel3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel3.Location = new System.Drawing.Point(407, 208);
            this.linkLabel3.Name = "linkLabel3";
            this.linkLabel3.Size = new System.Drawing.Size(446, 16);
            this.linkLabel3.TabIndex = 77;
            this.linkLabel3.TabStop = true;
            this.linkLabel3.Tag = "http://msdn.microsoft.com/en-us/library/aa168346%28v=office.11%29.aspx";
            this.linkLabel3.Text = "http://msdn.microsoft.com/en-us/library/aa168346%28v=office.11%29.aspx";
            this.linksToolTip.SetToolTip(this.linkLabel3, "http://msdn.microsoft.com/en-us/library/aa168346%28v=office.11%29.aspx");
            this.linkLabel3.Click += new System.EventHandler(this.linkLabel_Clicked);
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(407, 58);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(16, 16);
            this.pictureBox2.TabIndex = 79;
            this.pictureBox2.TabStop = false;
            // 
            // labelDidYouKnow
            // 
            this.labelDidYouKnow.AutoSize = true;
            this.labelDidYouKnow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.labelDidYouKnow.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelDidYouKnow.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelDidYouKnow.ForeColor = System.Drawing.Color.Black;
            this.labelDidYouKnow.Location = new System.Drawing.Point(426, 59);
            this.labelDidYouKnow.Name = "labelDidYouKnow";
            this.labelDidYouKnow.Size = new System.Drawing.Size(124, 16);
            this.labelDidYouKnow.TabIndex = 78;
            this.labelDidYouKnow.Text = "Did you know... ?";
            // 
            // labelSupressThisDialog
            // 
            this.labelSupressThisDialog.BackColor = System.Drawing.Color.Transparent;
            this.labelSupressThisDialog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelSupressThisDialog.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSupressThisDialog.ForeColor = System.Drawing.Color.Red;
            this.labelSupressThisDialog.Location = new System.Drawing.Point(140, 61);
            this.labelSupressThisDialog.Name = "labelSupressThisDialog";
            this.labelSupressThisDialog.Size = new System.Drawing.Size(213, 19);
            this.labelSupressThisDialog.TabIndex = 80;
            this.labelSupressThisDialog.Text = "Suppress this Dialog";
            // 
            // OutlookSecurityControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
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
            this.Controls.Add(this.labelMessages);
            this.Controls.Add(this.labelLog);
            this.Controls.Add(this.checkBoxSupressEnabled);
            this.Name = "OutlookSecurityControl";
            this.Size = new System.Drawing.Size(924, 496);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.LinkContextMenu.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkBoxSupressEnabled;
        private System.Windows.Forms.TextBox labelMessages;
        private System.Windows.Forms.Label labelLog;
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
        private System.Windows.Forms.ContextMenuStrip LinkContextMenu;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolTip linksToolTip;
    }
}
