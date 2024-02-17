namespace Sample.ExcelAddin
{
    partial class TranslationPane
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TranslationPane));
            this.comboBoxSourceLanguage = new System.Windows.Forms.ComboBox();
            this.labelSourceLanguage = new System.Windows.Forms.Label();
            this.labelTargetLanguage = new System.Windows.Forms.Label();
            this.comboBoxTargetLanguage = new System.Windows.Forms.ComboBox();
            this.textBoxRequested = new System.Windows.Forms.TextBox();
            this.labelRequestedText = new System.Windows.Forms.Label();
            this.labelTranslation = new System.Windows.Forms.Label();
            this.textBoxTranslation = new System.Windows.Forms.TextBox();
            this.checkBoxAutoTranslate = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.panelError = new System.Windows.Forms.Panel();
            this.pictureBoxInitial = new System.Windows.Forms.PictureBox();
            this.labelErrorMessage = new System.Windows.Forms.Label();
            this.pictureBoxError = new System.Windows.Forms.PictureBox();
            this.buttonTranslate = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.buttonDoLocalConnect = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.linkLabelNetOfficePage = new System.Windows.Forms.LinkLabel();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.panelError.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxInitial)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxError)).BeginInit();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBoxSourceLanguage
            // 
            this.comboBoxSourceLanguage.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboBoxSourceLanguage.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBoxSourceLanguage.FormattingEnabled = true;
            this.comboBoxSourceLanguage.Location = new System.Drawing.Point(61, 29);
            this.comboBoxSourceLanguage.Name = "comboBoxSourceLanguage";
            this.comboBoxSourceLanguage.Size = new System.Drawing.Size(127, 24);
            this.comboBoxSourceLanguage.TabIndex = 1;
            // 
            // labelSourceLanguage
            // 
            this.labelSourceLanguage.AutoSize = true;
            this.labelSourceLanguage.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSourceLanguage.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelSourceLanguage.Location = new System.Drawing.Point(16, 34);
            this.labelSourceLanguage.Name = "labelSourceLanguage";
            this.labelSourceLanguage.Size = new System.Drawing.Size(40, 16);
            this.labelSourceLanguage.TabIndex = 2;
            this.labelSourceLanguage.Text = "From";
            // 
            // labelTargetLanguage
            // 
            this.labelTargetLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelTargetLanguage.AutoSize = true;
            this.labelTargetLanguage.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTargetLanguage.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelTargetLanguage.Location = new System.Drawing.Point(26, 74);
            this.labelTargetLanguage.Name = "labelTargetLanguage";
            this.labelTargetLanguage.Size = new System.Drawing.Size(24, 16);
            this.labelTargetLanguage.TabIndex = 4;
            this.labelTargetLanguage.Text = "To";
            // 
            // comboBoxTargetLanguage
            // 
            this.comboBoxTargetLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.comboBoxTargetLanguage.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboBoxTargetLanguage.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBoxTargetLanguage.FormattingEnabled = true;
            this.comboBoxTargetLanguage.Location = new System.Drawing.Point(61, 69);
            this.comboBoxTargetLanguage.Name = "comboBoxTargetLanguage";
            this.comboBoxTargetLanguage.Size = new System.Drawing.Size(127, 24);
            this.comboBoxTargetLanguage.TabIndex = 3;
            // 
            // textBoxRequested
            // 
            this.textBoxRequested.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.textBoxRequested.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxRequested.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxRequested.Location = new System.Drawing.Point(17, 26);
            this.textBoxRequested.Multiline = true;
            this.textBoxRequested.Name = "textBoxRequested";
            this.textBoxRequested.Size = new System.Drawing.Size(325, 70);
            this.textBoxRequested.TabIndex = 5;
            this.textBoxRequested.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBoxRequested_KeyPress);
            // 
            // labelRequestedText
            // 
            this.labelRequestedText.AutoSize = true;
            this.labelRequestedText.BackColor = System.Drawing.Color.Transparent;
            this.labelRequestedText.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelRequestedText.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelRequestedText.Location = new System.Drawing.Point(16, 9);
            this.labelRequestedText.Name = "labelRequestedText";
            this.labelRequestedText.Size = new System.Drawing.Size(37, 16);
            this.labelRequestedText.TabIndex = 6;
            this.labelRequestedText.Text = "Text";
            // 
            // labelTranslation
            // 
            this.labelTranslation.AutoSize = true;
            this.labelTranslation.BackColor = System.Drawing.Color.Transparent;
            this.labelTranslation.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTranslation.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelTranslation.Location = new System.Drawing.Point(368, 9);
            this.labelTranslation.Name = "labelTranslation";
            this.labelTranslation.Size = new System.Drawing.Size(80, 16);
            this.labelTranslation.TabIndex = 8;
            this.labelTranslation.Text = "Translation";
            // 
            // textBoxTranslation
            // 
            this.textBoxTranslation.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxTranslation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxTranslation.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxTranslation.Location = new System.Drawing.Point(369, 26);
            this.textBoxTranslation.Multiline = true;
            this.textBoxTranslation.Name = "textBoxTranslation";
            this.textBoxTranslation.Size = new System.Drawing.Size(380, 70);
            this.textBoxTranslation.TabIndex = 7;
            // 
            // checkBoxAutoTranslate
            // 
            this.checkBoxAutoTranslate.AutoSize = true;
            this.checkBoxAutoTranslate.Checked = true;
            this.checkBoxAutoTranslate.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxAutoTranslate.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxAutoTranslate.ForeColor = System.Drawing.Color.MidnightBlue;
            this.checkBoxAutoTranslate.Location = new System.Drawing.Point(12, 29);
            this.checkBoxAutoTranslate.Name = "checkBoxAutoTranslate";
            this.checkBoxAutoTranslate.Size = new System.Drawing.Size(190, 20);
            this.checkBoxAutoTranslate.TabIndex = 9;
            this.checkBoxAutoTranslate.Text = "Auto Translate Selection";
            this.checkBoxAutoTranslate.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.labelTargetLanguage);
            this.groupBox1.Controls.Add(this.comboBoxTargetLanguage);
            this.groupBox1.Controls.Add(this.comboBoxSourceLanguage);
            this.groupBox1.Controls.Add(this.labelSourceLanguage);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(986, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(215, 131);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.panelError);
            this.groupBox2.Controls.Add(this.textBoxRequested);
            this.groupBox2.Controls.Add(this.buttonTranslate);
            this.groupBox2.Controls.Add(this.textBoxTranslation);
            this.groupBox2.Controls.Add(this.labelRequestedText);
            this.groupBox2.Controls.Add(this.labelTranslation);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(215, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(770, 131);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            // 
            // panelError
            // 
            this.panelError.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelError.Controls.Add(this.pictureBoxInitial);
            this.panelError.Controls.Add(this.labelErrorMessage);
            this.panelError.Controls.Add(this.pictureBoxError);
            this.panelError.Location = new System.Drawing.Point(369, 99);
            this.panelError.Name = "panelError";
            this.panelError.Size = new System.Drawing.Size(380, 23);
            this.panelError.TabIndex = 10;
            // 
            // pictureBoxInitial
            // 
            this.pictureBoxInitial.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBoxInitial.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxInitial.Image")));
            this.pictureBoxInitial.Location = new System.Drawing.Point(356, 2);
            this.pictureBoxInitial.Name = "pictureBoxInitial";
            this.pictureBoxInitial.Size = new System.Drawing.Size(21, 18);
            this.pictureBoxInitial.TabIndex = 14;
            this.pictureBoxInitial.TabStop = false;
            this.pictureBoxInitial.Visible = false;
            // 
            // labelErrorMessage
            // 
            this.labelErrorMessage.AutoSize = true;
            this.labelErrorMessage.ForeColor = System.Drawing.Color.Green;
            this.labelErrorMessage.Location = new System.Drawing.Point(31, 3);
            this.labelErrorMessage.Name = "labelErrorMessage";
            this.labelErrorMessage.Size = new System.Drawing.Size(265, 16);
            this.labelErrorMessage.TabIndex = 13;
            this.labelErrorMessage.Text = "Insert any text and translate with single click.";
            // 
            // pictureBoxError
            // 
            this.pictureBoxError.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxError.Image")));
            this.pictureBoxError.Location = new System.Drawing.Point(5, 3);
            this.pictureBoxError.Name = "pictureBoxError";
            this.pictureBoxError.Size = new System.Drawing.Size(21, 18);
            this.pictureBoxError.TabIndex = 12;
            this.pictureBoxError.TabStop = false;
            // 
            // buttonTranslate
            // 
            this.buttonTranslate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonTranslate.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonTranslate.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonTranslate.ForeColor = System.Drawing.Color.MidnightBlue;
            this.buttonTranslate.Image = ((System.Drawing.Image)(resources.GetObject("buttonTranslate.Image")));
            this.buttonTranslate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonTranslate.Location = new System.Drawing.Point(17, 96);
            this.buttonTranslate.Name = "buttonTranslate";
            this.buttonTranslate.Size = new System.Drawing.Size(325, 27);
            this.buttonTranslate.TabIndex = 9;
            this.buttonTranslate.Text = "Translate";
            this.buttonTranslate.UseVisualStyleBackColor = true;
            this.buttonTranslate.Click += new System.EventHandler(this.buttonTranslate_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox3.Controls.Add(this.buttonDoLocalConnect);
            this.groupBox3.Controls.Add(this.pictureBox1);
            this.groupBox3.Controls.Add(this.linkLabelNetOfficePage);
            this.groupBox3.Controls.Add(this.checkBoxAutoTranslate);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.groupBox3.Location = new System.Drawing.Point(0, 0);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(215, 131);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            // 
            // buttonDoLocalConnect
            // 
            this.buttonDoLocalConnect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonDoLocalConnect.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonDoLocalConnect.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonDoLocalConnect.ForeColor = System.Drawing.Color.MidnightBlue;
            this.buttonDoLocalConnect.Image = ((System.Drawing.Image)(resources.GetObject("buttonDoLocalConnect.Image")));
            this.buttonDoLocalConnect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonDoLocalConnect.Location = new System.Drawing.Point(12, 96);
            this.buttonDoLocalConnect.Name = "buttonDoLocalConnect";
            this.buttonDoLocalConnect.Size = new System.Drawing.Size(187, 27);
            this.buttonDoLocalConnect.TabIndex = 12;
            this.buttonDoLocalConnect.Text = "Connect again";
            this.buttonDoLocalConnect.UseVisualStyleBackColor = true;
            this.buttonDoLocalConnect.Click += new System.EventHandler(this.buttonDoLocalConnect_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(18, 57);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(33, 29);
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // linkLabelNetOfficePage
            // 
            this.linkLabelNetOfficePage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.linkLabelNetOfficePage.AutoSize = true;
            this.linkLabelNetOfficePage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelNetOfficePage.Location = new System.Drawing.Point(57, 64);
            this.linkLabelNetOfficePage.Name = "linkLabelNetOfficePage";
            this.linkLabelNetOfficePage.Size = new System.Drawing.Size(146, 16);
            this.linkLabelNetOfficePage.TabIndex = 10;
            this.linkLabelNetOfficePage.TabStop = true;
            this.linkLabelNetOfficePage.Text = "netoffice.codeplex.com";
            this.linkLabelNetOfficePage.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelNetOfficePage_LinkClicked);
            // 
            // TranslationPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "TranslationPane";
            this.Size = new System.Drawing.Size(1204, 130);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.panelError.ResumeLayout(false);
            this.panelError.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxInitial)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxError)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxSourceLanguage;
        private System.Windows.Forms.Label labelSourceLanguage;
        private System.Windows.Forms.Label labelTargetLanguage;
        private System.Windows.Forms.ComboBox comboBoxTargetLanguage;
        private System.Windows.Forms.TextBox textBoxRequested;
        private System.Windows.Forms.Label labelRequestedText;
        private System.Windows.Forms.Label labelTranslation;
        private System.Windows.Forms.TextBox textBoxTranslation;
        private System.Windows.Forms.CheckBox checkBoxAutoTranslate;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button buttonTranslate;
        private System.Windows.Forms.LinkLabel linkLabelNetOfficePage;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panelError;
        private System.Windows.Forms.PictureBox pictureBoxError;
        private System.Windows.Forms.Label labelErrorMessage;
        private System.Windows.Forms.PictureBox pictureBoxInitial;
        private System.Windows.Forms.Button buttonDoLocalConnect;
    }
}
