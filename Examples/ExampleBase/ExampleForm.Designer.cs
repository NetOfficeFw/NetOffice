namespace ExampleBase
{
    partial class ExampleForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExampleForm));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.FormImageList = new System.Windows.Forms.ImageList(this.components);
            this.HeaderLabel = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.OptionButton = new System.Windows.Forms.Button();
            this.LinkPanel = new System.Windows.Forms.Panel();
            this.Down1Label = new System.Windows.Forms.Label();
            this.Down2Label = new System.Windows.Forms.Label();
            this.LoveCodeplexPictureBox = new System.Windows.Forms.PictureBox();
            this.linkLabelCodeplex = new System.Windows.Forms.LinkLabel();
            this.linkLabelExampleOverview = new System.Windows.Forms.LinkLabel();
            this.linkLabelDiscussions = new System.Windows.Forms.LinkLabel();
            this.linkLabelGithub = new System.Windows.Forms.LinkLabel();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.linkLabelTutorialOverview = new System.Windows.Forms.LinkLabel();
            this.labelRessourceHeader = new System.Windows.Forms.Label();
            this.linkLabelDocumentation = new System.Windows.Forms.LinkLabel();
            this.ExamplesView = new System.Windows.Forms.DataGridView();
            this.IconColumn = new System.Windows.Forms.DataGridViewImageColumn();
            this.NameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DescriptionColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FormToolTip = new System.Windows.Forms.ToolTip(this.components);
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.LinkPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LoveCodeplexPictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ExamplesView)).BeginInit();
            this.SuspendLayout();
            // 
            // FormImageList
            // 
            this.FormImageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("FormImageList.ImageStream")));
            this.FormImageList.TransparentColor = System.Drawing.Color.Transparent;
            this.FormImageList.Images.SetKeyName(0, "example.png");
            // 
            // HeaderLabel
            // 
            this.HeaderLabel.AutoSize = true;
            this.HeaderLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.HeaderLabel.Location = new System.Drawing.Point(58, 25);
            this.HeaderLabel.Name = "HeaderLabel";
            this.HeaderLabel.Size = new System.Drawing.Size(228, 20);
            this.HeaderLabel.TabIndex = 1;
            this.HeaderLabel.Text = "Double click to run an example.";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(36, 28);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(18, 16);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // OptionButton
            // 
            this.OptionButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.OptionButton.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.OptionButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OptionButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.OptionButton.ForeColor = System.Drawing.Color.Blue;
            this.OptionButton.Image = ((System.Drawing.Image)(resources.GetObject("OptionButton.Image")));
            this.OptionButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.OptionButton.Location = new System.Drawing.Point(781, 22);
            this.OptionButton.Name = "OptionButton";
            this.OptionButton.Size = new System.Drawing.Size(154, 29);
            this.OptionButton.TabIndex = 12;
            this.OptionButton.Text = "  Options";
            this.OptionButton.UseVisualStyleBackColor = true;
            this.OptionButton.Click += new System.EventHandler(this.OptionButton_Click);
            // 
            // LinkPanel
            // 
            this.LinkPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LinkPanel.BackColor = System.Drawing.Color.White;
            this.LinkPanel.Controls.Add(this.linkLabel1);
            this.LinkPanel.Controls.Add(this.Down1Label);
            this.LinkPanel.Controls.Add(this.Down2Label);
            this.LinkPanel.Controls.Add(this.LoveCodeplexPictureBox);
            this.LinkPanel.Controls.Add(this.linkLabelCodeplex);
            this.LinkPanel.Controls.Add(this.linkLabelExampleOverview);
            this.LinkPanel.Controls.Add(this.linkLabelDiscussions);
            this.LinkPanel.Controls.Add(this.linkLabelGithub);
            this.LinkPanel.Controls.Add(this.pictureBox3);
            this.LinkPanel.Controls.Add(this.linkLabelTutorialOverview);
            this.LinkPanel.Controls.Add(this.labelRessourceHeader);
            this.LinkPanel.Controls.Add(this.linkLabelDocumentation);
            this.LinkPanel.Location = new System.Drawing.Point(781, 71);
            this.LinkPanel.Name = "LinkPanel";
            this.LinkPanel.Size = new System.Drawing.Size(154, 458);
            this.LinkPanel.TabIndex = 23;
            // 
            // Down1Label
            // 
            this.Down1Label.AutoSize = true;
            this.Down1Label.ForeColor = System.Drawing.Color.Gray;
            this.Down1Label.Location = new System.Drawing.Point(13, 281);
            this.Down1Label.Name = "Down1Label";
            this.Down1Label.Size = new System.Drawing.Size(89, 13);
            this.Down1Label.TabIndex = 35;
            this.Down1Label.Text = "actually read only";
            // 
            // Down2Label
            // 
            this.Down2Label.AutoSize = true;
            this.Down2Label.ForeColor = System.Drawing.Color.Gray;
            this.Down2Label.Location = new System.Drawing.Point(13, 300);
            this.Down2Label.Name = "Down2Label";
            this.Down2Label.Size = new System.Drawing.Size(114, 13);
            this.Down2Label.TabIndex = 34;
            this.Down2Label.Text = "shutdown 12/15/2017";
            // 
            // LoveCodeplexPictureBox
            // 
            this.LoveCodeplexPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("LoveCodeplexPictureBox.Image")));
            this.LoveCodeplexPictureBox.Location = new System.Drawing.Point(126, 258);
            this.LoveCodeplexPictureBox.Name = "LoveCodeplexPictureBox";
            this.LoveCodeplexPictureBox.Size = new System.Drawing.Size(19, 15);
            this.LoveCodeplexPictureBox.TabIndex = 33;
            this.LoveCodeplexPictureBox.TabStop = false;
            // 
            // linkLabelCodeplex
            // 
            this.linkLabelCodeplex.AutoSize = true;
            this.linkLabelCodeplex.Location = new System.Drawing.Point(10, 260);
            this.linkLabelCodeplex.Name = "linkLabelCodeplex";
            this.linkLabelCodeplex.Size = new System.Drawing.Size(114, 13);
            this.linkLabelCodeplex.TabIndex = 32;
            this.linkLabelCodeplex.TabStop = true;
            this.linkLabelCodeplex.Tag = "http://netoffice.codeplex.com";
            this.linkLabelCodeplex.Text = "NetOffice on Codeplex";
            this.FormToolTip.SetToolTip(this.linkLabelCodeplex, "http://netoffice.codeplex.com");
            this.linkLabelCodeplex.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.TopicLabel_LinkClicked);
            // 
            // linkLabelExampleOverview
            // 
            this.linkLabelExampleOverview.AutoSize = true;
            this.linkLabelExampleOverview.Location = new System.Drawing.Point(10, 190);
            this.linkLabelExampleOverview.Name = "linkLabelExampleOverview";
            this.linkLabelExampleOverview.Size = new System.Drawing.Size(133, 13);
            this.linkLabelExampleOverview.TabIndex = 31;
            this.linkLabelExampleOverview.TabStop = true;
            this.linkLabelExampleOverview.Tag = "https://netoffice.io/documentation/Examples.html";
            this.linkLabelExampleOverview.Text = "Examples Online Overview";
            this.FormToolTip.SetToolTip(this.linkLabelExampleOverview, "https://netoffice.io/documentation/Examples.html");
            this.linkLabelExampleOverview.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.TopicLabel_LinkClicked);
            // 
            // linkLabelDiscussions
            // 
            this.linkLabelDiscussions.AutoSize = true;
            this.linkLabelDiscussions.Location = new System.Drawing.Point(10, 130);
            this.linkLabelDiscussions.Name = "linkLabelDiscussions";
            this.linkLabelDiscussions.Size = new System.Drawing.Size(89, 13);
            this.linkLabelDiscussions.TabIndex = 26;
            this.linkLabelDiscussions.TabStop = true;
            this.linkLabelDiscussions.Tag = "https://github.com/NetOfficeFw/NetOffice/issues";
            this.linkLabelDiscussions.Text = "Discussion Board";
            this.FormToolTip.SetToolTip(this.linkLabelDiscussions, "https://github.com/NetOfficeFw/NetOffice/issues");
            this.linkLabelDiscussions.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.TopicLabel_LinkClicked);
            // 
            // linkLabelGithub
            // 
            this.linkLabelGithub.AutoSize = true;
            this.linkLabelGithub.Location = new System.Drawing.Point(10, 100);
            this.linkLabelGithub.Name = "linkLabelGithub";
            this.linkLabelGithub.Size = new System.Drawing.Size(101, 13);
            this.linkLabelGithub.TabIndex = 24;
            this.linkLabelGithub.TabStop = true;
            this.linkLabelGithub.Tag = "https://github.com/NetOfficeFw";
            this.linkLabelGithub.Text = "NetOffice on Github";
            this.FormToolTip.SetToolTip(this.linkLabelGithub, "https://github.com/NetOfficeFw");
            this.linkLabelGithub.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.TopicLabel_LinkClicked);
            // 
            // pictureBox3
            // 
            this.pictureBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
            this.pictureBox3.Location = new System.Drawing.Point(108, 7);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(32, 32);
            this.pictureBox3.TabIndex = 23;
            this.pictureBox3.TabStop = false;
            // 
            // linkLabelTutorialOverview
            // 
            this.linkLabelTutorialOverview.AutoSize = true;
            this.linkLabelTutorialOverview.Location = new System.Drawing.Point(10, 160);
            this.linkLabelTutorialOverview.Name = "linkLabelTutorialOverview";
            this.linkLabelTutorialOverview.Size = new System.Drawing.Size(128, 13);
            this.linkLabelTutorialOverview.TabIndex = 19;
            this.linkLabelTutorialOverview.TabStop = true;
            this.linkLabelTutorialOverview.Tag = "https://netoffice.io/documentation/Tutorials.html";
            this.linkLabelTutorialOverview.Text = "Tutorials Online Overview";
            this.FormToolTip.SetToolTip(this.linkLabelTutorialOverview, "https://netoffice.io/documentation/Tutorials.html");
            this.linkLabelTutorialOverview.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.TopicLabel_LinkClicked);
            // 
            // labelRessourceHeader
            // 
            this.labelRessourceHeader.AutoSize = true;
            this.labelRessourceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelRessourceHeader.Location = new System.Drawing.Point(10, 12);
            this.labelRessourceHeader.Name = "labelRessourceHeader";
            this.labelRessourceHeader.Size = new System.Drawing.Size(67, 13);
            this.labelRessourceHeader.TabIndex = 18;
            this.labelRessourceHeader.Text = "Resources";
            // 
            // linkLabelDocumentation
            // 
            this.linkLabelDocumentation.AutoSize = true;
            this.linkLabelDocumentation.Location = new System.Drawing.Point(10, 220);
            this.linkLabelDocumentation.Name = "linkLabelDocumentation";
            this.linkLabelDocumentation.Size = new System.Drawing.Size(79, 13);
            this.linkLabelDocumentation.TabIndex = 15;
            this.linkLabelDocumentation.TabStop = true;
            this.linkLabelDocumentation.Tag = "https://netoffice.io/documentation/index.html";
            this.linkLabelDocumentation.Text = "Documentation";
            this.FormToolTip.SetToolTip(this.linkLabelDocumentation, "https://netoffice.io/documentation/index.html");
            this.linkLabelDocumentation.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.TopicLabel_LinkClicked);
            // 
            // ExamplesView
            // 
            this.ExamplesView.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.GhostWhite;
            this.ExamplesView.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.ExamplesView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ExamplesView.BackgroundColor = System.Drawing.Color.White;
            this.ExamplesView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.ExamplesView.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleHorizontal;
            this.ExamplesView.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.ExamplesView.ColumnHeadersHeight = 30;
            this.ExamplesView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.IconColumn,
            this.NameColumn,
            this.DescriptionColumn});
            this.ExamplesView.Cursor = System.Windows.Forms.Cursors.Hand;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.Orange;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.ExamplesView.DefaultCellStyle = dataGridViewCellStyle2;
            this.ExamplesView.Location = new System.Drawing.Point(33, 71);
            this.ExamplesView.MultiSelect = false;
            this.ExamplesView.Name = "ExamplesView";
            this.ExamplesView.RowHeadersVisible = false;
            this.ExamplesView.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.ExamplesView.RowTemplate.Height = 32;
            this.ExamplesView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.ExamplesView.Size = new System.Drawing.Size(716, 458);
            this.ExamplesView.TabIndex = 24;
            this.ExamplesView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.ExamplesView_MouseDoubleClick);
            // 
            // IconColumn
            // 
            this.IconColumn.DataPropertyName = "Icon";
            this.IconColumn.HeaderText = "";
            this.IconColumn.Name = "IconColumn";
            this.IconColumn.Width = 40;
            // 
            // NameColumn
            // 
            this.NameColumn.DataPropertyName = "Caption";
            this.NameColumn.HeaderText = "Name";
            this.NameColumn.Name = "NameColumn";
            this.NameColumn.Width = 160;
            // 
            // DescriptionColumn
            // 
            this.DescriptionColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.DescriptionColumn.DataPropertyName = "Description";
            this.DescriptionColumn.HeaderText = "Description";
            this.DescriptionColumn.Name = "DescriptionColumn";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(10, 70);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(95, 13);
            this.linkLabel1.TabIndex = 36;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Tag = "https://osdn.net/projects/netoffice";
            this.linkLabel1.Text = "NetOffice on Osdn";
            this.FormToolTip.SetToolTip(this.linkLabel1, "https://osdn.net/projects/netoffice");
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.TopicLabel_LinkClicked);
            // 
            // ExampleForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.ClientSize = new System.Drawing.Size(966, 562);
            this.Controls.Add(this.ExamplesView);
            this.Controls.Add(this.LinkPanel);
            this.Controls.Add(this.OptionButton);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.HeaderLabel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(982, 600);
            this.Name = "ExampleForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormBase";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.LinkPanel.ResumeLayout(false);
            this.LinkPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LoveCodeplexPictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ExamplesView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label HeaderLabel;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ImageList FormImageList;
        private System.Windows.Forms.Button OptionButton;
        private System.Windows.Forms.Panel LinkPanel;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.LinkLabel linkLabelTutorialOverview;
        private System.Windows.Forms.Label labelRessourceHeader;
        private System.Windows.Forms.LinkLabel linkLabelDocumentation;
        private System.Windows.Forms.DataGridView ExamplesView;
        private System.Windows.Forms.DataGridViewImageColumn IconColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn NameColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn DescriptionColumn;
        private System.Windows.Forms.LinkLabel linkLabelGithub;
        private System.Windows.Forms.LinkLabel linkLabelDiscussions;
        private System.Windows.Forms.ToolTip FormToolTip;
        private System.Windows.Forms.LinkLabel linkLabelExampleOverview;
        private System.Windows.Forms.Label Down1Label;
        private System.Windows.Forms.Label Down2Label;
        private System.Windows.Forms.PictureBox LoveCodeplexPictureBox;
        private System.Windows.Forms.LinkLabel linkLabelCodeplex;
        private System.Windows.Forms.LinkLabel linkLabel1;
    }
}

