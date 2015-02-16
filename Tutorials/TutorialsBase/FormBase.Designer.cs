namespace TutorialsBase
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
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.panelTutorials = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panelShowTutorialLink = new System.Windows.Forms.Panel();
            this.labelOffHint = new System.Windows.Forms.Label();
            this.linkLabelTutorialContent = new System.Windows.Forms.LinkLabel();
            this.webBrowserTutorialContent = new System.Windows.Forms.WebBrowser();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panelTutorialArea = new System.Windows.Forms.Panel();
            this.buttonRunTutorial = new System.Windows.Forms.Button();
            this.linkLabelDiscussionBoard = new System.Windows.Forms.LinkLabel();
            this.labelQuestions = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelHeader2 = new System.Windows.Forms.Label();
            this.buttonOptions = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.linkLabelAccess = new System.Windows.Forms.LinkLabel();
            this.linkLabelDeveloperToolbox = new System.Windows.Forms.LinkLabel();
            this.linkLabelTutorialOverview = new System.Windows.Forms.LinkLabel();
            this.linkLabelPowerPoint = new System.Windows.Forms.LinkLabel();
            this.linkLabelTecFaq = new System.Windows.Forms.LinkLabel();
            this.labelRessourceHeader = new System.Windows.Forms.Label();
            this.linkLabelOutlook = new System.Windows.Forms.LinkLabel();
            this.linkLabelTecDocumentation = new System.Windows.Forms.LinkLabel();
            this.linkLabelExcel = new System.Windows.Forms.LinkLabel();
            this.linkLabelWord = new System.Windows.Forms.LinkLabel();
            this.linkLabelDocumentation = new System.Windows.Forms.LinkLabel();
            this.labelTutorialDescription = new System.Windows.Forms.Label();
            this.listViewTutorials = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.panelTutorials.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panelShowTutorialLink.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panelTutorialArea.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "example.png");
            // 
            // panelTutorials
            // 
            this.panelTutorials.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelTutorials.BackColor = System.Drawing.SystemColors.Control;
            this.panelTutorials.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelTutorials.Controls.Add(this.tabControl1);
            this.panelTutorials.Location = new System.Drawing.Point(105, 37);
            this.panelTutorials.Name = "panelTutorials";
            this.panelTutorials.Size = new System.Drawing.Size(702, 512);
            this.panelTutorials.TabIndex = 3;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(700, 510);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.panelShowTutorialLink);
            this.tabPage1.Controls.Add(this.webBrowserTutorialContent);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(692, 484);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Introduction";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // panelShowTutorialLink
            // 
            this.panelShowTutorialLink.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelShowTutorialLink.BackColor = System.Drawing.Color.LightSteelBlue;
            this.panelShowTutorialLink.Controls.Add(this.labelOffHint);
            this.panelShowTutorialLink.Controls.Add(this.linkLabelTutorialContent);
            this.panelShowTutorialLink.Location = new System.Drawing.Point(0, 0);
            this.panelShowTutorialLink.Name = "panelShowTutorialLink";
            this.panelShowTutorialLink.Size = new System.Drawing.Size(690, 171);
            this.panelShowTutorialLink.TabIndex = 7;
            this.panelShowTutorialLink.Visible = false;
            // 
            // labelOffHint
            // 
            this.labelOffHint.AutoSize = true;
            this.labelOffHint.Location = new System.Drawing.Point(25, 18);
            this.labelOffHint.Name = "labelOffHint";
            this.labelOffHint.Size = new System.Drawing.Size(334, 13);
            this.labelOffHint.TabIndex = 7;
            this.labelOffHint.Text = "Online Documentation is turned off. Enable in program options or visit:";
            // 
            // linkLabelTutorialContent
            // 
            this.linkLabelTutorialContent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelTutorialContent.AutoSize = true;
            this.linkLabelTutorialContent.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelTutorialContent.Location = new System.Drawing.Point(51, 42);
            this.linkLabelTutorialContent.Name = "linkLabelTutorialContent";
            this.linkLabelTutorialContent.Size = new System.Drawing.Size(152, 16);
            this.linkLabelTutorialContent.TabIndex = 6;
            this.linkLabelTutorialContent.TabStop = true;
            this.linkLabelTutorialContent.Tag = "http://netoffice.codeplex.com";
            this.linkLabelTutorialContent.Text = "linkLabelTutorialContent";
            this.linkLabelTutorialContent.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelTutorialContent_LinkClicked);
            // 
            // webBrowserTutorialContent
            // 
            this.webBrowserTutorialContent.AllowWebBrowserDrop = false;
            this.webBrowserTutorialContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowserTutorialContent.IsWebBrowserContextMenuEnabled = false;
            this.webBrowserTutorialContent.Location = new System.Drawing.Point(3, 3);
            this.webBrowserTutorialContent.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowserTutorialContent.Name = "webBrowserTutorialContent";
            this.webBrowserTutorialContent.ScriptErrorsSuppressed = true;
            this.webBrowserTutorialContent.Size = new System.Drawing.Size(686, 478);
            this.webBrowserTutorialContent.TabIndex = 0;
            this.webBrowserTutorialContent.WebBrowserShortcutsEnabled = false;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.panelTutorialArea);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(692, 484);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Code Sample";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // panelTutorialArea
            // 
            this.panelTutorialArea.BackColor = System.Drawing.Color.LightSteelBlue;
            this.panelTutorialArea.Controls.Add(this.buttonRunTutorial);
            this.panelTutorialArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelTutorialArea.Location = new System.Drawing.Point(3, 3);
            this.panelTutorialArea.Name = "panelTutorialArea";
            this.panelTutorialArea.Size = new System.Drawing.Size(686, 478);
            this.panelTutorialArea.TabIndex = 0;
            // 
            // buttonRunTutorial
            // 
            this.buttonRunTutorial.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonRunTutorial.Font = new System.Drawing.Font("MS Reference Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRunTutorial.Image = ((System.Drawing.Image)(resources.GetObject("buttonRunTutorial.Image")));
            this.buttonRunTutorial.Location = new System.Drawing.Point(194, 89);
            this.buttonRunTutorial.Name = "buttonRunTutorial";
            this.buttonRunTutorial.Size = new System.Drawing.Size(274, 285);
            this.buttonRunTutorial.TabIndex = 0;
            this.buttonRunTutorial.Text = "Click here to run Example";
            this.buttonRunTutorial.UseVisualStyleBackColor = true;
            this.buttonRunTutorial.Visible = false;
            this.buttonRunTutorial.Click += new System.EventHandler(this.buttonRunTutorial_Click);
            // 
            // linkLabelDiscussionBoard
            // 
            this.linkLabelDiscussionBoard.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelDiscussionBoard.AutoSize = true;
            this.linkLabelDiscussionBoard.Location = new System.Drawing.Point(665, 562);
            this.linkLabelDiscussionBoard.Name = "linkLabelDiscussionBoard";
            this.linkLabelDiscussionBoard.Size = new System.Drawing.Size(137, 13);
            this.linkLabelDiscussionBoard.TabIndex = 5;
            this.linkLabelDiscussionBoard.TabStop = true;
            this.linkLabelDiscussionBoard.Tag = "http://netoffice.codeplex.com/discussions";
            this.linkLabelDiscussionBoard.Text = "NetOffice Discussion Board";
            this.linkLabelDiscussionBoard.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelDiscussion_LinkClicked);
            // 
            // labelQuestions
            // 
            this.labelQuestions.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.labelQuestions.AutoSize = true;
            this.labelQuestions.Location = new System.Drawing.Point(535, 562);
            this.labelQuestions.Name = "labelQuestions";
            this.labelQuestions.Size = new System.Drawing.Size(124, 13);
            this.labelQuestions.TabIndex = 6;
            this.labelQuestions.Text = "Questions or Comments?";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(139, 560);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(18, 16);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // labelHeader2
            // 
            this.labelHeader2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelHeader2.AutoSize = true;
            this.labelHeader2.Location = new System.Drawing.Point(159, 562);
            this.labelHeader2.Name = "labelHeader2";
            this.labelHeader2.Size = new System.Drawing.Size(273, 13);
            this.labelHeader2.TabIndex = 11;
            this.labelHeader2.Text = "The tutorial code is available in the VS Solution Explorer.";
            // 
            // buttonOptions
            // 
            this.buttonOptions.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOptions.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOptions.Image = ((System.Drawing.Image)(resources.GetObject("buttonOptions.Image")));
            this.buttonOptions.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonOptions.Location = new System.Drawing.Point(813, 519);
            this.buttonOptions.Name = "buttonOptions";
            this.buttonOptions.Size = new System.Drawing.Size(154, 29);
            this.buttonOptions.TabIndex = 12;
            this.buttonOptions.Text = "Options";
            this.buttonOptions.UseVisualStyleBackColor = true;
            this.buttonOptions.Click += new System.EventHandler(this.buttonOptions_Click);
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BackColor = System.Drawing.Color.White;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.pictureBox2);
            this.panel2.Controls.Add(this.linkLabelAccess);
            this.panel2.Controls.Add(this.linkLabelDeveloperToolbox);
            this.panel2.Controls.Add(this.linkLabelTutorialOverview);
            this.panel2.Controls.Add(this.linkLabelPowerPoint);
            this.panel2.Controls.Add(this.linkLabelTecFaq);
            this.panel2.Controls.Add(this.labelRessourceHeader);
            this.panel2.Controls.Add(this.linkLabelOutlook);
            this.panel2.Controls.Add(this.linkLabelTecDocumentation);
            this.panel2.Controls.Add(this.linkLabelExcel);
            this.panel2.Controls.Add(this.linkLabelWord);
            this.panel2.Controls.Add(this.linkLabelDocumentation);
            this.panel2.Location = new System.Drawing.Point(813, 37);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(154, 476);
            this.panel2.TabIndex = 22;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(106, 7);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(32, 32);
            this.pictureBox2.TabIndex = 23;
            this.pictureBox2.TabStop = false;
            // 
            // linkLabelAccess
            // 
            this.linkLabelAccess.AutoSize = true;
            this.linkLabelAccess.Location = new System.Drawing.Point(10, 256);
            this.linkLabelAccess.Name = "linkLabelAccess";
            this.linkLabelAccess.Size = new System.Drawing.Size(90, 13);
            this.linkLabelAccess.TabIndex = 21;
            this.linkLabelAccess.TabStop = true;
            this.linkLabelAccess.Tag = "/wikipage?title=Access_Examples_EN#/wikipage?title=Access_Examples_DE";
            this.linkLabelAccess.Text = "Access Examples";
            this.linkLabelAccess.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // linkLabelDeveloperToolbox
            // 
            this.linkLabelDeveloperToolbox.AutoSize = true;
            this.linkLabelDeveloperToolbox.Location = new System.Drawing.Point(10, 138);
            this.linkLabelDeveloperToolbox.Name = "linkLabelDeveloperToolbox";
            this.linkLabelDeveloperToolbox.Size = new System.Drawing.Size(97, 13);
            this.linkLabelDeveloperToolbox.TabIndex = 20;
            this.linkLabelDeveloperToolbox.TabStop = true;
            this.linkLabelDeveloperToolbox.Tag = "/wikipage?title=DeveloperToolbox_English#/wikipage?title=DeveloperToolbox_German";
            this.linkLabelDeveloperToolbox.Text = "Developer Toolbox";
            this.linkLabelDeveloperToolbox.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // linkLabelTutorialOverview
            // 
            this.linkLabelTutorialOverview.AutoSize = true;
            this.linkLabelTutorialOverview.Location = new System.Drawing.Point(10, 46);
            this.linkLabelTutorialOverview.Name = "linkLabelTutorialOverview";
            this.linkLabelTutorialOverview.Size = new System.Drawing.Size(90, 13);
            this.linkLabelTutorialOverview.TabIndex = 19;
            this.linkLabelTutorialOverview.TabStop = true;
            this.linkLabelTutorialOverview.Tag = "/wikipage?title=TutorialOverview_EN#/wikipage?title=TutorialOverview_DE";
            this.linkLabelTutorialOverview.Text = "Tutorial Overview";
            this.linkLabelTutorialOverview.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // linkLabelPowerPoint
            // 
            this.linkLabelPowerPoint.AutoSize = true;
            this.linkLabelPowerPoint.Location = new System.Drawing.Point(10, 233);
            this.linkLabelPowerPoint.Name = "linkLabelPowerPoint";
            this.linkLabelPowerPoint.Size = new System.Drawing.Size(109, 13);
            this.linkLabelPowerPoint.TabIndex = 14;
            this.linkLabelPowerPoint.TabStop = true;
            this.linkLabelPowerPoint.Tag = "/wikipage?title=PPoint_Examples_EN#/wikipage?title=PPoint_Examples_DE";
            this.linkLabelPowerPoint.Text = "PowerPoint Examples";
            this.linkLabelPowerPoint.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // linkLabelTecFaq
            // 
            this.linkLabelTecFaq.AutoSize = true;
            this.linkLabelTecFaq.Location = new System.Drawing.Point(10, 115);
            this.linkLabelTecFaq.Name = "linkLabelTecFaq";
            this.linkLabelTecFaq.Size = new System.Drawing.Size(78, 13);
            this.linkLabelTecFaq.TabIndex = 14;
            this.linkLabelTecFaq.TabStop = true;
            this.linkLabelTecFaq.Tag = "/wikipage?title=Tec_Faq_English#/wikipage?title=Tec_Faq_German";
            this.linkLabelTecFaq.Text = "Technical FAQ";
            this.linkLabelTecFaq.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
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
            // linkLabelOutlook
            // 
            this.linkLabelOutlook.AutoSize = true;
            this.linkLabelOutlook.Location = new System.Drawing.Point(10, 210);
            this.linkLabelOutlook.Name = "linkLabelOutlook";
            this.linkLabelOutlook.Size = new System.Drawing.Size(92, 13);
            this.linkLabelOutlook.TabIndex = 13;
            this.linkLabelOutlook.TabStop = true;
            this.linkLabelOutlook.Tag = "/wikipage?title=Outlook_Examples_EN#/wikipage?title=Outlook_Examples_DE";
            this.linkLabelOutlook.Text = "Outlook Examples";
            this.linkLabelOutlook.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // linkLabelTecDocumentation
            // 
            this.linkLabelTecDocumentation.AutoSize = true;
            this.linkLabelTecDocumentation.Location = new System.Drawing.Point(10, 92);
            this.linkLabelTecDocumentation.Name = "linkLabelTecDocumentation";
            this.linkLabelTecDocumentation.Size = new System.Drawing.Size(129, 13);
            this.linkLabelTecDocumentation.TabIndex = 13;
            this.linkLabelTecDocumentation.TabStop = true;
            this.linkLabelTecDocumentation.Tag = "/wikipage?title=Tec_Documentation_English#/wikipage?title=Tec_Documentation_Germa" +
                "n";
            this.linkLabelTecDocumentation.Text = "Technical Documentation";
            this.linkLabelTecDocumentation.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // linkLabelExcel
            // 
            this.linkLabelExcel.AutoSize = true;
            this.linkLabelExcel.Location = new System.Drawing.Point(10, 163);
            this.linkLabelExcel.Name = "linkLabelExcel";
            this.linkLabelExcel.Size = new System.Drawing.Size(81, 13);
            this.linkLabelExcel.TabIndex = 15;
            this.linkLabelExcel.TabStop = true;
            this.linkLabelExcel.Tag = "/wikipage?title=Excel_Examples_EN#/wikipage?title=Excel_Examples_DE";
            this.linkLabelExcel.Text = "Excel Examples";
            this.linkLabelExcel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // linkLabelWord
            // 
            this.linkLabelWord.AutoSize = true;
            this.linkLabelWord.Location = new System.Drawing.Point(10, 187);
            this.linkLabelWord.Name = "linkLabelWord";
            this.linkLabelWord.Size = new System.Drawing.Size(81, 13);
            this.linkLabelWord.TabIndex = 16;
            this.linkLabelWord.TabStop = true;
            this.linkLabelWord.Tag = "/wikipage?title=Word_Examples_EN#/wikipage?title=Word_Examples_DE";
            this.linkLabelWord.Text = "Word Examples";
            this.linkLabelWord.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // linkLabelDocumentation
            // 
            this.linkLabelDocumentation.AutoSize = true;
            this.linkLabelDocumentation.Location = new System.Drawing.Point(10, 69);
            this.linkLabelDocumentation.Name = "linkLabelDocumentation";
            this.linkLabelDocumentation.Size = new System.Drawing.Size(79, 13);
            this.linkLabelDocumentation.TabIndex = 15;
            this.linkLabelDocumentation.TabStop = true;
            this.linkLabelDocumentation.Tag = "/documentation#/wikipage?title=Documentation_German";
            this.linkLabelDocumentation.Text = "Documentation";
            this.linkLabelDocumentation.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMultiLanguage_LinkClicked);
            // 
            // labelTutorialDescription
            // 
            this.labelTutorialDescription.AutoSize = true;
            this.labelTutorialDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTutorialDescription.Location = new System.Drawing.Point(110, 14);
            this.labelTutorialDescription.Name = "labelTutorialDescription";
            this.labelTutorialDescription.Size = new System.Drawing.Size(175, 16);
            this.labelTutorialDescription.TabIndex = 23;
            this.labelTutorialDescription.Text = "labelTutorialDescription";
            // 
            // listViewTutorials
            // 
            this.listViewTutorials.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.listViewTutorials.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listViewTutorials.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.listViewTutorials.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listViewTutorials.Location = new System.Drawing.Point(6, 37);
            this.listViewTutorials.MultiSelect = false;
            this.listViewTutorials.Name = "listViewTutorials";
            this.listViewTutorials.Size = new System.Drawing.Size(95, 512);
            this.listViewTutorials.SmallImageList = this.imageList1;
            this.listViewTutorials.TabIndex = 26;
            this.listViewTutorials.UseCompatibleStateImageBehavior = false;
            this.listViewTutorials.View = System.Windows.Forms.View.Details;
            this.listViewTutorials.SelectedIndexChanged += new System.EventHandler(this.listViewTutorials_SelectedIndexChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Tutorials";
            this.columnHeader1.Width = 75;
            // 
            // FormBase
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(973, 584);
            this.Controls.Add(this.listViewTutorials);
            this.Controls.Add(this.labelTutorialDescription);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.buttonOptions);
            this.Controls.Add(this.labelHeader2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.labelQuestions);
            this.Controls.Add(this.linkLabelDiscussionBoard);
            this.Controls.Add(this.panelTutorials);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormBase";
            this.Text = "FormBase";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormBase_FormClosed);
            this.Resize += new System.EventHandler(this.FormBase_Resize);
            this.panelTutorials.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panelShowTutorialLink.ResumeLayout(false);
            this.panelShowTutorialLink.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.panelTutorialArea.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panelTutorials;
        private System.Windows.Forms.LinkLabel linkLabelDiscussionBoard;
        private System.Windows.Forms.Label labelQuestions;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label labelHeader2;
        private System.Windows.Forms.Button buttonOptions;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.LinkLabel linkLabelTecFaq;
        private System.Windows.Forms.Label labelRessourceHeader;
        private System.Windows.Forms.LinkLabel linkLabelTecDocumentation;
        private System.Windows.Forms.LinkLabel linkLabelDocumentation;
        private System.Windows.Forms.LinkLabel linkLabelPowerPoint;
        private System.Windows.Forms.LinkLabel linkLabelOutlook;
        private System.Windows.Forms.LinkLabel linkLabelExcel;
        private System.Windows.Forms.LinkLabel linkLabelWord;
        private System.Windows.Forms.LinkLabel linkLabelTutorialOverview;
        private System.Windows.Forms.LinkLabel linkLabelDeveloperToolbox;
        private System.Windows.Forms.LinkLabel linkLabelAccess;
        private System.Windows.Forms.Label labelTutorialDescription;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.WebBrowser webBrowserTutorialContent;
        private System.Windows.Forms.ListView listViewTutorials;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.Panel panelTutorialArea;
        private System.Windows.Forms.LinkLabel linkLabelTutorialContent;
        private System.Windows.Forms.Button buttonRunTutorial;
        private System.Windows.Forms.Panel panelShowTutorialLink;
        private System.Windows.Forms.Label labelOffHint;
        private System.Windows.Forms.PictureBox pictureBox2;
    }
}

