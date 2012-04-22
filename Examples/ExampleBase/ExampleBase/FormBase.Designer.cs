namespace ExampleBase
{
    partial class FormBase
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormBase));
            this.listViewExamples = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.labelHeader1 = new System.Windows.Forms.Label();
            this.buttonStartExample = new System.Windows.Forms.Button();
            this.panelExamples = new System.Windows.Forms.Panel();
            this.labelControlsHeader = new System.Windows.Forms.Label();
            this.linkLabelDiscussionBoard = new System.Windows.Forms.LinkLabel();
            this.labelQuestions = new System.Windows.Forms.Label();
            this.labelWeWantYou = new System.Windows.Forms.Label();
            this.linkLabelEmployeWanted = new System.Windows.Forms.LinkLabel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.labelHeader2 = new System.Windows.Forms.Label();
            this.buttonOptions = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.linkLabelProjectWizard = new System.Windows.Forms.LinkLabel();
            this.linkLabelDeveloperToolbox = new System.Windows.Forms.LinkLabel();
            this.linkLabelTutorialOverview = new System.Windows.Forms.LinkLabel();
            this.linkLabelTecFaq = new System.Windows.Forms.LinkLabel();
            this.labelRessourceHeader = new System.Windows.Forms.Label();
            this.linkLabelTecDocumentation = new System.Windows.Forms.LinkLabel();
            this.linkLabelDocumentation = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            this.SuspendLayout();
            // 
            // listViewExamples
            // 
            this.listViewExamples.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewExamples.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this.listViewExamples.Cursor = System.Windows.Forms.Cursors.Hand;
            this.listViewExamples.FullRowSelect = true;
            this.listViewExamples.HideSelection = false;
            this.listViewExamples.Location = new System.Drawing.Point(24, 61);
            this.listViewExamples.Name = "listViewExamples";
            this.listViewExamples.Size = new System.Drawing.Size(715, 187);
            this.listViewExamples.SmallImageList = this.imageList1;
            this.listViewExamples.TabIndex = 0;
            this.listViewExamples.UseCompatibleStateImageBehavior = false;
            this.listViewExamples.View = System.Windows.Forms.View.Details;
            this.listViewExamples.SelectedIndexChanged += new System.EventHandler(this.listViewExamples_SelectedIndexChanged);
            this.listViewExamples.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listViewExamples_MouseDoubleClick);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Name";
            this.columnHeader1.Width = 93;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Description";
            this.columnHeader2.Width = 587;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "example.png");
            // 
            // labelHeader1
            // 
            this.labelHeader1.AutoSize = true;
            this.labelHeader1.Location = new System.Drawing.Point(48, 19);
            this.labelHeader1.Name = "labelHeader1";
            this.labelHeader1.Size = new System.Drawing.Size(402, 13);
            this.labelHeader1.TabIndex = 1;
            this.labelHeader1.Text = "Click on the button \"Start Example\" or double click in the list view to run an ex" +
                "ample";
            // 
            // buttonStartExample
            // 
            this.buttonStartExample.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonStartExample.Image = ((System.Drawing.Image)(resources.GetObject("buttonStartExample.Image")));
            this.buttonStartExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonStartExample.Location = new System.Drawing.Point(619, 18);
            this.buttonStartExample.Name = "buttonStartExample";
            this.buttonStartExample.Size = new System.Drawing.Size(120, 29);
            this.buttonStartExample.TabIndex = 2;
            this.buttonStartExample.Text = "Start Example";
            this.buttonStartExample.UseVisualStyleBackColor = true;
            this.buttonStartExample.Click += new System.EventHandler(this.buttonStartExample_Click);
            // 
            // panelExamples
            // 
            this.panelExamples.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelExamples.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelExamples.Location = new System.Drawing.Point(24, 308);
            this.panelExamples.Name = "panelExamples";
            this.panelExamples.Size = new System.Drawing.Size(715, 304);
            this.panelExamples.TabIndex = 3;
            // 
            // labelControlsHeader
            // 
            this.labelControlsHeader.AutoSize = true;
            this.labelControlsHeader.Location = new System.Drawing.Point(48, 278);
            this.labelControlsHeader.Name = "labelControlsHeader";
            this.labelControlsHeader.Size = new System.Drawing.Size(248, 13);
            this.labelControlsHeader.TabIndex = 4;
            this.labelControlsHeader.Text = "This area is for examples with an own visual control";
            // 
            // linkLabelDiscussionBoard
            // 
            this.linkLabelDiscussionBoard.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelDiscussionBoard.AutoSize = true;
            this.linkLabelDiscussionBoard.Location = new System.Drawing.Point(597, 628);
            this.linkLabelDiscussionBoard.Name = "linkLabelDiscussionBoard";
            this.linkLabelDiscussionBoard.Size = new System.Drawing.Size(137, 13);
            this.linkLabelDiscussionBoard.TabIndex = 5;
            this.linkLabelDiscussionBoard.TabStop = true;
            this.linkLabelDiscussionBoard.Tag = "http://netoffice.codeplex.com/discussions";
            this.linkLabelDiscussionBoard.Text = "NetOffice Discussion Board";
            this.linkLabelDiscussionBoard.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelDiscussionBoard_LinkClicked);
            // 
            // labelQuestions
            // 
            this.labelQuestions.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.labelQuestions.AutoSize = true;
            this.labelQuestions.Location = new System.Drawing.Point(467, 628);
            this.labelQuestions.Name = "labelQuestions";
            this.labelQuestions.Size = new System.Drawing.Size(124, 13);
            this.labelQuestions.TabIndex = 6;
            this.labelQuestions.Text = "Questions or Comments?";
            // 
            // labelWeWantYou
            // 
            this.labelWeWantYou.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.labelWeWantYou.AutoSize = true;
            this.labelWeWantYou.Location = new System.Drawing.Point(30, 628);
            this.labelWeWantYou.Name = "labelWeWantYou";
            this.labelWeWantYou.Size = new System.Drawing.Size(276, 13);
            this.labelWeWantYou.TabIndex = 7;
            this.labelWeWantYou.Text = "NetOffice is looking for developers/editors/office experts.";
            // 
            // linkLabelEmployeWanted
            // 
            this.linkLabelEmployeWanted.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelEmployeWanted.AutoSize = true;
            this.linkLabelEmployeWanted.Location = new System.Drawing.Point(312, 628);
            this.linkLabelEmployeWanted.Name = "linkLabelEmployeWanted";
            this.linkLabelEmployeWanted.Size = new System.Drawing.Size(103, 13);
            this.linkLabelEmployeWanted.TabIndex = 8;
            this.linkLabelEmployeWanted.TabStop = true;
            this.linkLabelEmployeWanted.Tag = "http://netoffice.codeplex.com/wikipage?title=JoinNetOffice_English";
            this.linkLabelEmployeWanted.Text = "NetOffice Job Board";
            this.linkLabelEmployeWanted.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelEmployeWanted_LinkClicked);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(28, 18);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(18, 16);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(29, 277);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(18, 16);
            this.pictureBox2.TabIndex = 10;
            this.pictureBox2.TabStop = false;
            // 
            // labelHeader2
            // 
            this.labelHeader2.AutoSize = true;
            this.labelHeader2.Location = new System.Drawing.Point(48, 38);
            this.labelHeader2.Name = "labelHeader2";
            this.labelHeader2.Size = new System.Drawing.Size(281, 13);
            this.labelHeader2.TabIndex = 11;
            this.labelHeader2.Text = "The example code is available in the VS Solution Explorer.";
            // 
            // buttonOptions
            // 
            this.buttonOptions.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOptions.Image = ((System.Drawing.Image)(resources.GetObject("buttonOptions.Image")));
            this.buttonOptions.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonOptions.Location = new System.Drawing.Point(517, 18);
            this.buttonOptions.Name = "buttonOptions";
            this.buttonOptions.Size = new System.Drawing.Size(94, 29);
            this.buttonOptions.TabIndex = 12;
            this.buttonOptions.Text = "Options";
            this.buttonOptions.UseVisualStyleBackColor = true;
            this.buttonOptions.Click += new System.EventHandler(this.buttonOptions_Click);
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.pictureBox3);
            this.panel2.Controls.Add(this.linkLabelProjectWizard);
            this.panel2.Controls.Add(this.linkLabelDeveloperToolbox);
            this.panel2.Controls.Add(this.linkLabelTutorialOverview);
            this.panel2.Controls.Add(this.linkLabelTecFaq);
            this.panel2.Controls.Add(this.labelRessourceHeader);
            this.panel2.Controls.Add(this.linkLabelTecDocumentation);
            this.panel2.Controls.Add(this.linkLabelDocumentation);
            this.panel2.Location = new System.Drawing.Point(763, 61);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(154, 551);
            this.panel2.TabIndex = 23;
            // 
            // pictureBox3
            // 
            this.pictureBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
            this.pictureBox3.Location = new System.Drawing.Point(106, 7);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(32, 32);
            this.pictureBox3.TabIndex = 23;
            this.pictureBox3.TabStop = false;
            // 
            // linkLabelProjectWizard
            // 
            this.linkLabelProjectWizard.AutoSize = true;
            this.linkLabelProjectWizard.Location = new System.Drawing.Point(10, 140);
            this.linkLabelProjectWizard.Name = "linkLabelProjectWizard";
            this.linkLabelProjectWizard.Size = new System.Drawing.Size(140, 13);
            this.linkLabelProjectWizard.TabIndex = 22;
            this.linkLabelProjectWizard.TabStop = true;
            this.linkLabelProjectWizard.Tag = "/wikipage?title=ProjectWizardScreenshots_English#/wikipage?title=ProjectWizardScr" +
                "eenshots_German";
            this.linkLabelProjectWizard.Text = "Visual Studio Project Wizard";
            this.linkLabelProjectWizard.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelRessouce_LinkClicked);
            // 
            // linkLabelDeveloperToolbox
            // 
            this.linkLabelDeveloperToolbox.AutoSize = true;
            this.linkLabelDeveloperToolbox.Location = new System.Drawing.Point(10, 120);
            this.linkLabelDeveloperToolbox.Name = "linkLabelDeveloperToolbox";
            this.linkLabelDeveloperToolbox.Size = new System.Drawing.Size(97, 13);
            this.linkLabelDeveloperToolbox.TabIndex = 20;
            this.linkLabelDeveloperToolbox.TabStop = true;
            this.linkLabelDeveloperToolbox.Tag = "/wikipage?title=DeveloperToolbox_English#/wikipage?title=DeveloperToolbox_German";
            this.linkLabelDeveloperToolbox.Text = "Developer Toolbox";
            this.linkLabelDeveloperToolbox.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelRessouce_LinkClicked);
            // 
            // linkLabelTutorialOverview
            // 
            this.linkLabelTutorialOverview.AutoSize = true;
            this.linkLabelTutorialOverview.Location = new System.Drawing.Point(10, 40);
            this.linkLabelTutorialOverview.Name = "linkLabelTutorialOverview";
            this.linkLabelTutorialOverview.Size = new System.Drawing.Size(47, 13);
            this.linkLabelTutorialOverview.TabIndex = 19;
            this.linkLabelTutorialOverview.TabStop = true;
            this.linkLabelTutorialOverview.Tag = "/wikipage?title=TutorialOverview_EN#/wikipage?title=TutorialOverview_DE";
            this.linkLabelTutorialOverview.Text = "Tutorials";
            this.linkLabelTutorialOverview.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelRessouce_LinkClicked);
            // 
            // linkLabelTecFaq
            // 
            this.linkLabelTecFaq.AutoSize = true;
            this.linkLabelTecFaq.Location = new System.Drawing.Point(10, 100);
            this.linkLabelTecFaq.Name = "linkLabelTecFaq";
            this.linkLabelTecFaq.Size = new System.Drawing.Size(78, 13);
            this.linkLabelTecFaq.TabIndex = 14;
            this.linkLabelTecFaq.TabStop = true;
            this.linkLabelTecFaq.Tag = "/wikipage?title=Tec_Faq_English#/wikipage?title=Tec_Faq_German";
            this.linkLabelTecFaq.Text = "Technical FAQ";
            this.linkLabelTecFaq.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelRessouce_LinkClicked);
            // 
            // labelRessourceHeader
            // 
            this.labelRessourceHeader.AutoSize = true;
            this.labelRessourceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelRessourceHeader.Location = new System.Drawing.Point(10, 12);
            this.labelRessourceHeader.Name = "labelRessourceHeader";
            this.labelRessourceHeader.Size = new System.Drawing.Size(73, 13);
            this.labelRessourceHeader.TabIndex = 18;
            this.labelRessourceHeader.Text = "Ressources";
            // 
            // linkLabelTecDocumentation
            // 
            this.linkLabelTecDocumentation.AutoSize = true;
            this.linkLabelTecDocumentation.Location = new System.Drawing.Point(10, 80);
            this.linkLabelTecDocumentation.Name = "linkLabelTecDocumentation";
            this.linkLabelTecDocumentation.Size = new System.Drawing.Size(129, 13);
            this.linkLabelTecDocumentation.TabIndex = 13;
            this.linkLabelTecDocumentation.TabStop = true;
            this.linkLabelTecDocumentation.Tag = "/wikipage?title=Tec_Documentation_English#/wikipage?title=Tec_Documentation_Germa" +
                "n";
            this.linkLabelTecDocumentation.Text = "Technical Documentation";
            this.linkLabelTecDocumentation.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelRessouce_LinkClicked);
            // 
            // linkLabelDocumentation
            // 
            this.linkLabelDocumentation.AutoSize = true;
            this.linkLabelDocumentation.Location = new System.Drawing.Point(10, 60);
            this.linkLabelDocumentation.Name = "linkLabelDocumentation";
            this.linkLabelDocumentation.Size = new System.Drawing.Size(79, 13);
            this.linkLabelDocumentation.TabIndex = 15;
            this.linkLabelDocumentation.TabStop = true;
            this.linkLabelDocumentation.Tag = "/documentation#/wikipage?title=Documentation_German";
            this.linkLabelDocumentation.Text = "Documentation";
            this.linkLabelDocumentation.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelRessouce_LinkClicked);
            // 
            // FormBase
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(940, 655);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.buttonOptions);
            this.Controls.Add(this.labelHeader2);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.linkLabelEmployeWanted);
            this.Controls.Add(this.labelWeWantYou);
            this.Controls.Add(this.labelQuestions);
            this.Controls.Add(this.linkLabelDiscussionBoard);
            this.Controls.Add(this.labelControlsHeader);
            this.Controls.Add(this.panelExamples);
            this.Controls.Add(this.buttonStartExample);
            this.Controls.Add(this.labelHeader1);
            this.Controls.Add(this.listViewExamples);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FormBase";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormBase";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView listViewExamples;
        private System.Windows.Forms.Label labelHeader1;
        private System.Windows.Forms.Button buttonStartExample;
        private System.Windows.Forms.Panel panelExamples;
        private System.Windows.Forms.Label labelControlsHeader;
        private System.Windows.Forms.LinkLabel linkLabelDiscussionBoard;
        private System.Windows.Forms.Label labelQuestions;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.Label labelWeWantYou;
        private System.Windows.Forms.LinkLabel linkLabelEmployeWanted;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label labelHeader2;
        private System.Windows.Forms.Button buttonOptions;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.LinkLabel linkLabelProjectWizard;
        private System.Windows.Forms.LinkLabel linkLabelDeveloperToolbox;
        private System.Windows.Forms.LinkLabel linkLabelTutorialOverview;
        private System.Windows.Forms.LinkLabel linkLabelTecFaq;
        private System.Windows.Forms.Label labelRessourceHeader;
        private System.Windows.Forms.LinkLabel linkLabelTecDocumentation;
        private System.Windows.Forms.LinkLabel linkLabelDocumentation;
    }
}

