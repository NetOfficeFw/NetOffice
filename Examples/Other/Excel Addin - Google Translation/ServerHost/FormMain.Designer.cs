namespace Sample.ServerHost
{
    partial class FormMain
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

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.pictureBoxPermission = new System.Windows.Forms.PictureBox();
            this.labelPermission = new System.Windows.Forms.Label();
            this.buttonClose = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.buttonTranslate = new System.Windows.Forms.Button();
            this.textBoxRequested = new System.Windows.Forms.TextBox();
            this.textBoxTranslation = new System.Windows.Forms.TextBox();
            this.labelRequestedText = new System.Windows.Forms.Label();
            this.labelTranslation = new System.Windows.Forms.Label();
            this.labelTargetLanguage = new System.Windows.Forms.Label();
            this.comboBoxTargetLanguage = new System.Windows.Forms.ComboBox();
            this.comboBoxSourceLanguage = new System.Windows.Forms.ComboBox();
            this.labelSourceLanguage = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxPermission)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "start.png");
            this.imageList1.Images.SetKeyName(1, "ok.png");
            this.imageList1.Images.SetKeyName(2, "warning.png");
            this.imageList1.Images.SetKeyName(3, "error.png");
            this.imageList1.Images.SetKeyName(4, "color_wheel.png");
            this.imageList1.Images.SetKeyName(5, "data_refresh.png");
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.listView1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.pictureBoxPermission);
            this.splitContainer1.Panel2.Controls.Add(this.labelPermission);
            this.splitContainer1.Panel2.Controls.Add(this.buttonClose);
            this.splitContainer1.Size = new System.Drawing.Size(721, 292);
            this.splitContainer1.SplitterDistance = 235;
            this.splitContainer1.TabIndex = 0;
            // 
            // listView1
            // 
            this.listView1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this.listView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listView1.FullRowSelect = true;
            this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listView1.Location = new System.Drawing.Point(0, 0);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(721, 235);
            this.listView1.SmallImageList = this.imageList1;
            this.listView1.TabIndex = 0;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "";
            this.columnHeader1.Width = 166;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "";
            this.columnHeader2.Width = 500;
            // 
            // pictureBoxPermission
            // 
            this.pictureBoxPermission.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxPermission.Image")));
            this.pictureBoxPermission.Location = new System.Drawing.Point(22, 19);
            this.pictureBoxPermission.Name = "pictureBoxPermission";
            this.pictureBoxPermission.Size = new System.Drawing.Size(18, 18);
            this.pictureBoxPermission.TabIndex = 3;
            this.pictureBoxPermission.TabStop = false;
            // 
            // labelPermission
            // 
            this.labelPermission.AutoSize = true;
            this.labelPermission.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelPermission.Location = new System.Drawing.Point(43, 21);
            this.labelPermission.Name = "labelPermission";
            this.labelPermission.Size = new System.Drawing.Size(359, 13);
            this.labelPermission.TabIndex = 2;
            this.labelPermission.Text = "The application want establish a HTTP connection to Google Translations.";
            this.labelPermission.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // buttonClose
            // 
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonClose.ForeColor = System.Drawing.Color.MidnightBlue;
            this.buttonClose.Image = ((System.Drawing.Image)(resources.GetObject("buttonClose.Image")));
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.Location = new System.Drawing.Point(520, 13);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(137, 28);
            this.buttonClose.TabIndex = 0;
            this.buttonClose.Text = "Close Server";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.ImageList = this.imageList1;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(729, 322);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tabPage1.Controls.Add(this.splitContainer1);
            this.tabPage1.ImageIndex = 4;
            this.tabPage1.Location = new System.Drawing.Point(4, 26);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(721, 292);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Activity";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tabPage2.Controls.Add(this.buttonTranslate);
            this.tabPage2.Controls.Add(this.textBoxRequested);
            this.tabPage2.Controls.Add(this.textBoxTranslation);
            this.tabPage2.Controls.Add(this.labelRequestedText);
            this.tabPage2.Controls.Add(this.labelTranslation);
            this.tabPage2.Controls.Add(this.labelTargetLanguage);
            this.tabPage2.Controls.Add(this.comboBoxTargetLanguage);
            this.tabPage2.Controls.Add(this.comboBoxSourceLanguage);
            this.tabPage2.Controls.Add(this.labelSourceLanguage);
            this.tabPage2.ImageIndex = 5;
            this.tabPage2.Location = new System.Drawing.Point(4, 26);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(721, 292);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Translation";
            // 
            // buttonTranslate
            // 
            this.buttonTranslate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonTranslate.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonTranslate.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonTranslate.ForeColor = System.Drawing.Color.MidnightBlue;
            this.buttonTranslate.Image = ((System.Drawing.Image)(resources.GetObject("buttonTranslate.Image")));
            this.buttonTranslate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonTranslate.Location = new System.Drawing.Point(36, 144);
            this.buttonTranslate.Name = "buttonTranslate";
            this.buttonTranslate.Size = new System.Drawing.Size(648, 27);
            this.buttonTranslate.TabIndex = 13;
            this.buttonTranslate.Text = "Translate (Ctrl + Return)";
            this.buttonTranslate.UseVisualStyleBackColor = true;
            this.buttonTranslate.Click += new System.EventHandler(this.buttonTranslate_Click);
            // 
            // textBoxRequested
            // 
            this.textBoxRequested.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxRequested.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxRequested.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxRequested.Location = new System.Drawing.Point(36, 78);
            this.textBoxRequested.Multiline = true;
            this.textBoxRequested.Name = "textBoxRequested";
            this.textBoxRequested.Size = new System.Drawing.Size(648, 69);
            this.textBoxRequested.TabIndex = 9;
            this.textBoxRequested.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxRequested_KeyDown);
            // 
            // textBoxTranslation
            // 
            this.textBoxTranslation.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxTranslation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxTranslation.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxTranslation.Location = new System.Drawing.Point(36, 204);
            this.textBoxTranslation.Multiline = true;
            this.textBoxTranslation.Name = "textBoxTranslation";
            this.textBoxTranslation.Size = new System.Drawing.Size(648, 70);
            this.textBoxTranslation.TabIndex = 11;
            // 
            // labelRequestedText
            // 
            this.labelRequestedText.AutoSize = true;
            this.labelRequestedText.BackColor = System.Drawing.Color.Transparent;
            this.labelRequestedText.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelRequestedText.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelRequestedText.Location = new System.Drawing.Point(33, 59);
            this.labelRequestedText.Name = "labelRequestedText";
            this.labelRequestedText.Size = new System.Drawing.Size(37, 16);
            this.labelRequestedText.TabIndex = 10;
            this.labelRequestedText.Text = "Text";
            // 
            // labelTranslation
            // 
            this.labelTranslation.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelTranslation.AutoSize = true;
            this.labelTranslation.BackColor = System.Drawing.Color.Transparent;
            this.labelTranslation.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTranslation.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelTranslation.Location = new System.Drawing.Point(33, 185);
            this.labelTranslation.Name = "labelTranslation";
            this.labelTranslation.Size = new System.Drawing.Size(80, 16);
            this.labelTranslation.TabIndex = 12;
            this.labelTranslation.Text = "Translation";
            // 
            // labelTargetLanguage
            // 
            this.labelTargetLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelTargetLanguage.AutoSize = true;
            this.labelTargetLanguage.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTargetLanguage.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelTargetLanguage.Location = new System.Drawing.Point(414, 26);
            this.labelTargetLanguage.Name = "labelTargetLanguage";
            this.labelTargetLanguage.Size = new System.Drawing.Size(24, 16);
            this.labelTargetLanguage.TabIndex = 8;
            this.labelTargetLanguage.Text = "To";
            // 
            // comboBoxTargetLanguage
            // 
            this.comboBoxTargetLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxTargetLanguage.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboBoxTargetLanguage.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBoxTargetLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxTargetLanguage.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBoxTargetLanguage.FormattingEnabled = true;
            this.comboBoxTargetLanguage.Location = new System.Drawing.Point(448, 23);
            this.comboBoxTargetLanguage.Name = "comboBoxTargetLanguage";
            this.comboBoxTargetLanguage.Size = new System.Drawing.Size(235, 21);
            this.comboBoxTargetLanguage.TabIndex = 7;
            // 
            // comboBoxSourceLanguage
            // 
            this.comboBoxSourceLanguage.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboBoxSourceLanguage.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBoxSourceLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxSourceLanguage.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBoxSourceLanguage.FormattingEnabled = true;
            this.comboBoxSourceLanguage.Location = new System.Drawing.Point(78, 23);
            this.comboBoxSourceLanguage.Name = "comboBoxSourceLanguage";
            this.comboBoxSourceLanguage.Size = new System.Drawing.Size(235, 21);
            this.comboBoxSourceLanguage.TabIndex = 5;
            // 
            // labelSourceLanguage
            // 
            this.labelSourceLanguage.AutoSize = true;
            this.labelSourceLanguage.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSourceLanguage.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelSourceLanguage.Location = new System.Drawing.Point(33, 27);
            this.labelSourceLanguage.Name = "labelSourceLanguage";
            this.labelSourceLanguage.Size = new System.Drawing.Size(40, 16);
            this.labelSourceLanguage.TabIndex = 6;
            this.labelSourceLanguage.Text = "From";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.ClientSize = new System.Drawing.Size(729, 322);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sample.ServerHost - NetOffice Google Translation";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxPermission)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Label labelPermission;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label labelTargetLanguage;
        private System.Windows.Forms.ComboBox comboBoxTargetLanguage;
        private System.Windows.Forms.ComboBox comboBoxSourceLanguage;
        private System.Windows.Forms.Label labelSourceLanguage;
        private System.Windows.Forms.TextBox textBoxRequested;
        private System.Windows.Forms.TextBox textBoxTranslation;
        private System.Windows.Forms.Label labelRequestedText;
        private System.Windows.Forms.Label labelTranslation;
        private System.Windows.Forms.Button buttonTranslate;
        private System.Windows.Forms.PictureBox pictureBoxPermission;

    }
}

