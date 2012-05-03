namespace NetOffice.ProjectWizard
{
    partial class WizardDialog
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WizardDialog));
            this.nextButton = new System.Windows.Forms.Button();
            this.backButton = new System.Windows.Forms.Button();
            this.panelControls = new System.Windows.Forms.Panel();
            this.finishButton = new System.Windows.Forms.Button();
            this.labelCurrentStep = new System.Windows.Forms.Label();
            this.labelBottomDelimiter = new System.Windows.Forms.Label();
            this.labelDescription = new System.Windows.Forms.Label();
            this.labelCaption = new System.Windows.Forms.Label();
            this.imageBox = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.imageListIcons = new System.Windows.Forms.ImageList(this.components);
            this.cancelButton = new System.Windows.Forms.Button();
            this.comboBoxLanguage = new System.Windows.Forms.ComboBox();
            this.labelSelectLanguageHeader = new System.Windows.Forms.Label();
            this.panelLeftHeader = new System.Windows.Forms.Panel();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBoxNetOfficeLogo = new System.Windows.Forms.PictureBox();
            this.labelMiddleDelimiter = new System.Windows.Forms.Label();
            this.labelTopDelimiter = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.imageBox)).BeginInit();
            this.panel1.SuspendLayout();
            this.panelLeftHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxNetOfficeLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // nextButton
            // 
            this.nextButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.nextButton.Enabled = false;
            this.nextButton.Image = ((System.Drawing.Image)(resources.GetObject("nextButton.Image")));
            this.nextButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.nextButton.Location = new System.Drawing.Point(447, 290);
            this.nextButton.Name = "nextButton";
            this.nextButton.Size = new System.Drawing.Size(92, 23);
            this.nextButton.TabIndex = 24;
            this.nextButton.Text = "Weiter ";
            this.nextButton.Click += new System.EventHandler(this.nextButton_Click);
            // 
            // backButton
            // 
            this.backButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.backButton.Enabled = false;
            this.backButton.Image = ((System.Drawing.Image)(resources.GetObject("backButton.Image")));
            this.backButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.backButton.Location = new System.Drawing.Point(345, 290);
            this.backButton.Name = "backButton";
            this.backButton.Size = new System.Drawing.Size(92, 23);
            this.backButton.TabIndex = 25;
            this.backButton.Text = " Zurück";
            this.backButton.Click += new System.EventHandler(this.backButton_Click);
            // 
            // panelControls
            // 
            this.panelControls.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelControls.Location = new System.Drawing.Point(98, 60);
            this.panelControls.Name = "panelControls";
            this.panelControls.Size = new System.Drawing.Size(466, 214);
            this.panelControls.TabIndex = 26;
            // 
            // finishButton
            // 
            this.finishButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.finishButton.Image = ((System.Drawing.Image)(resources.GetObject("finishButton.Image")));
            this.finishButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.finishButton.Location = new System.Drawing.Point(324, 290);
            this.finishButton.Name = "finishButton";
            this.finishButton.Size = new System.Drawing.Size(92, 23);
            this.finishButton.TabIndex = 27;
            this.finishButton.Text = "Fertig  ";
            this.finishButton.Click += new System.EventHandler(this.finishButton_Click);
            // 
            // labelCurrentStep
            // 
            this.labelCurrentStep.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelCurrentStep.Location = new System.Drawing.Point(2, 247);
            this.labelCurrentStep.Name = "labelCurrentStep";
            this.labelCurrentStep.Size = new System.Drawing.Size(92, 13);
            this.labelCurrentStep.TabIndex = 28;
            this.labelCurrentStep.Text = "Schritt {0} von {1}";
            this.labelCurrentStep.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // labelBottomDelimiter
            // 
            this.labelBottomDelimiter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelBottomDelimiter.Location = new System.Drawing.Point(0, 273);
            this.labelBottomDelimiter.Name = "labelBottomDelimiter";
            this.labelBottomDelimiter.Size = new System.Drawing.Size(622, 1);
            this.labelBottomDelimiter.TabIndex = 29;
            // 
            // labelDescription
            // 
            this.labelDescription.AutoSize = true;
            this.labelDescription.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.labelDescription.Location = new System.Drawing.Point(30, 34);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(60, 13);
            this.labelDescription.TabIndex = 33;
            this.labelDescription.Text = "Description";
            // 
            // labelCaption
            // 
            this.labelCaption.AutoSize = true;
            this.labelCaption.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.labelCaption.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCaption.Location = new System.Drawing.Point(14, 13);
            this.labelCaption.Name = "labelCaption";
            this.labelCaption.Size = new System.Drawing.Size(50, 15);
            this.labelCaption.TabIndex = 31;
            this.labelCaption.Text = "Caption";
            // 
            // imageBox
            // 
            this.imageBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.imageBox.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.imageBox.Image = ((System.Drawing.Image)(resources.GetObject("imageBox.Image")));
            this.imageBox.Location = new System.Drawing.Point(411, 13);
            this.imageBox.Name = "imageBox";
            this.imageBox.Size = new System.Drawing.Size(34, 33);
            this.imageBox.TabIndex = 30;
            this.imageBox.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.panel1.Controls.Add(this.labelCaption);
            this.panel1.Controls.Add(this.imageBox);
            this.panel1.Controls.Add(this.labelDescription);
            this.panel1.Location = new System.Drawing.Point(98, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(483, 60);
            this.panel1.TabIndex = 34;
            // 
            // imageListIcons
            // 
            this.imageListIcons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListIcons.ImageStream")));
            this.imageListIcons.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListIcons.Images.SetKeyName(0, "question.png");
            this.imageListIcons.Images.SetKeyName(1, "information.png");
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.cancelButton.Image = ((System.Drawing.Image)(resources.GetObject("cancelButton.Image")));
            this.cancelButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.cancelButton.Location = new System.Drawing.Point(419, 290);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(92, 23);
            this.cancelButton.TabIndex = 35;
            this.cancelButton.Text = "      Abbrechen ";
            this.cancelButton.Visible = false;
            // 
            // comboBoxLanguage
            // 
            this.comboBoxLanguage.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.comboBoxLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxLanguage.FormattingEnabled = true;
            this.comboBoxLanguage.Items.AddRange(new object[] {
            "English",
            "German"});
            this.comboBoxLanguage.Location = new System.Drawing.Point(160, 292);
            this.comboBoxLanguage.Name = "comboBoxLanguage";
            this.comboBoxLanguage.Size = new System.Drawing.Size(102, 21);
            this.comboBoxLanguage.TabIndex = 36;
            this.comboBoxLanguage.SelectedIndexChanged += new System.EventHandler(this.comboBoxLanguage_SelectedIndexChanged);
            // 
            // labelSelectLanguageHeader
            // 
            this.labelSelectLanguageHeader.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.labelSelectLanguageHeader.AutoSize = true;
            this.labelSelectLanguageHeader.Location = new System.Drawing.Point(101, 295);
            this.labelSelectLanguageHeader.Name = "labelSelectLanguageHeader";
            this.labelSelectLanguageHeader.Size = new System.Drawing.Size(55, 13);
            this.labelSelectLanguageHeader.TabIndex = 37;
            this.labelSelectLanguageHeader.Text = "Language";
            // 
            // panelLeftHeader
            // 
            this.panelLeftHeader.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.panelLeftHeader.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panelLeftHeader.Controls.Add(this.pictureBox4);
            this.panelLeftHeader.Controls.Add(this.pictureBox3);
            this.panelLeftHeader.Controls.Add(this.pictureBox2);
            this.panelLeftHeader.Controls.Add(this.pictureBox1);
            this.panelLeftHeader.Controls.Add(this.pictureBoxNetOfficeLogo);
            this.panelLeftHeader.Controls.Add(this.labelCurrentStep);
            this.panelLeftHeader.Location = new System.Drawing.Point(0, 0);
            this.panelLeftHeader.Name = "panelLeftHeader";
            this.panelLeftHeader.Size = new System.Drawing.Size(98, 273);
            this.panelLeftHeader.TabIndex = 38;
            // 
            // pictureBox4
            // 
            this.pictureBox4.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox4.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox4.Image")));
            this.pictureBox4.Location = new System.Drawing.Point(30, 159);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(31, 32);
            this.pictureBox4.TabIndex = 34;
            this.pictureBox4.TabStop = false;
            // 
            // pictureBox3
            // 
            this.pictureBox3.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
            this.pictureBox3.Location = new System.Drawing.Point(31, 116);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(32, 32);
            this.pictureBox3.TabIndex = 33;
            this.pictureBox3.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(30, 72);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(31, 33);
            this.pictureBox2.TabIndex = 32;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(30, 202);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(33, 32);
            this.pictureBox1.TabIndex = 31;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBoxNetOfficeLogo
            // 
            this.pictureBoxNetOfficeLogo.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.pictureBoxNetOfficeLogo.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxNetOfficeLogo.Image")));
            this.pictureBoxNetOfficeLogo.Location = new System.Drawing.Point(23, 11);
            this.pictureBoxNetOfficeLogo.Name = "pictureBoxNetOfficeLogo";
            this.pictureBoxNetOfficeLogo.Size = new System.Drawing.Size(50, 43);
            this.pictureBoxNetOfficeLogo.TabIndex = 30;
            this.pictureBoxNetOfficeLogo.TabStop = false;
            // 
            // labelMiddleDelimiter
            // 
            this.labelMiddleDelimiter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelMiddleDelimiter.Location = new System.Drawing.Point(98, 60);
            this.labelMiddleDelimiter.Name = "labelMiddleDelimiter";
            this.labelMiddleDelimiter.Size = new System.Drawing.Size(524, 1);
            this.labelMiddleDelimiter.TabIndex = 39;
            // 
            // labelTopDelimiter
            // 
            this.labelTopDelimiter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelTopDelimiter.Location = new System.Drawing.Point(97, 1);
            this.labelTopDelimiter.Name = "labelTopDelimiter";
            this.labelTopDelimiter.Size = new System.Drawing.Size(524, 1);
            this.labelTopDelimiter.TabIndex = 40;
            // 
            // WizardDialog
            // 
            this.AcceptButton = this.nextButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(556, 319);
            this.ControlBox = false;
            this.Controls.Add(this.labelTopDelimiter);
            this.Controls.Add(this.labelMiddleDelimiter);
            this.Controls.Add(this.panelLeftHeader);
            this.Controls.Add(this.labelSelectLanguageHeader);
            this.Controls.Add(this.comboBoxLanguage);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.labelBottomDelimiter);
            this.Controls.Add(this.finishButton);
            this.Controls.Add(this.panelControls);
            this.Controls.Add(this.backButton);
            this.Controls.Add(this.nextButton);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WizardDialog";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "WizardDialog";
            ((System.ComponentModel.ISupportInitialize)(this.imageBox)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panelLeftHeader.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxNetOfficeLogo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button nextButton;
        private System.Windows.Forms.Button backButton;
        private System.Windows.Forms.Panel panelControls;
        private System.Windows.Forms.Button finishButton;
        private System.Windows.Forms.Label labelCurrentStep;
        private System.Windows.Forms.Label labelBottomDelimiter;
        private System.Windows.Forms.Label labelDescription;
        private System.Windows.Forms.Label labelCaption;
        private System.Windows.Forms.PictureBox imageBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ImageList imageListIcons;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.ComboBox comboBoxLanguage;
        private System.Windows.Forms.Label labelSelectLanguageHeader;
        private System.Windows.Forms.Panel panelLeftHeader;
        private System.Windows.Forms.PictureBox pictureBoxNetOfficeLogo;
        private System.Windows.Forms.Label labelMiddleDelimiter;
        private System.Windows.Forms.Label labelTopDelimiter;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.PictureBox pictureBox2;
    }
}
