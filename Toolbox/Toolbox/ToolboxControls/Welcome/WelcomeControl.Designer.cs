namespace NetOffice.DeveloperToolbox.ToolboxControls.Welcome
{
    partial class WelcomeControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WelcomeControl));
            this.labelVersionHint = new System.Windows.Forms.Label();
            this.panelMainLeft = new System.Windows.Forms.Panel();
            this.labelBeginTitle = new System.Windows.Forms.Label();
            this.pictureBoxIconLeft = new System.Windows.Forms.PictureBox();
            this.checkBoxStartAppWithWindows = new System.Windows.Forms.CheckBox();
            this.checkBoxStartAppMinimized = new System.Windows.Forms.CheckBox();
            this.checkBoxMinimizeToTray = new System.Windows.Forms.CheckBox();
            this.labelLanguage = new System.Windows.Forms.Label();
            this.comboBoxLanguage = new System.Windows.Forms.ComboBox();
            this.pictureBoxLogo = new System.Windows.Forms.PictureBox();
            this.panelMainRight = new System.Windows.Forms.Panel();
            this.labelMailMe = new System.Windows.Forms.Label();
            this.linkLabelMailMe = new System.Windows.Forms.LinkLabel();
            this.labelQuestion = new System.Windows.Forms.Label();
            this.linkLabelNetOfficeQuestions = new System.Windows.Forms.LinkLabel();
            this.labelUpdate = new System.Windows.Forms.Label();
            this.pictureBoxIconRight = new System.Windows.Forms.PictureBox();
            this.labelBug = new System.Windows.Forms.Label();
            this.linkLabelNetOfficeIssues = new System.Windows.Forms.LinkLabel();
            this.labelIWant = new System.Windows.Forms.Label();
            this.linkLabelNetOfficeUpdates = new System.Windows.Forms.LinkLabel();
            this.panelOptions = new System.Windows.Forms.Panel();
            this.buttonLanguageEditor = new System.Windows.Forms.Button();
            this.pictureBoxHeader = new System.Windows.Forms.PictureBox();
            this.labelBeginBottom = new NetOffice.DeveloperToolbox.Controls.Text.AdvRichTextBox();
            this.labelBeginTop = new NetOffice.DeveloperToolbox.Controls.Text.AdvRichTextBox();
            this.controlForeColorAnimator1 = new NetOffice.DeveloperToolbox.Utils.Animation.ControlForeColorAnimator(this.components);
            this.panelMainLeft.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxIconLeft)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLogo)).BeginInit();
            this.panelMainRight.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxIconRight)).BeginInit();
            this.panelOptions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.controlForeColorAnimator1)).BeginInit();
            this.SuspendLayout();
            // 
            // labelVersionHint
            // 
            this.labelVersionHint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.labelVersionHint.AutoSize = true;
            this.labelVersionHint.BackColor = System.Drawing.Color.Transparent;
            this.labelVersionHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelVersionHint.ForeColor = System.Drawing.Color.Gray;
            this.labelVersionHint.Location = new System.Drawing.Point(400, 381);
            this.labelVersionHint.Name = "labelVersionHint";
            this.labelVersionHint.Size = new System.Drawing.Size(123, 16);
            this.labelVersionHint.TabIndex = 101;
            this.labelVersionHint.Text = "labelVersionHint";
            this.labelVersionHint.Visible = false;
            // 
            // panelMainLeft
            // 
            this.panelMainLeft.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.panelMainLeft.Controls.Add(this.labelBeginBottom);
            this.panelMainLeft.Controls.Add(this.labelBeginTop);
            this.panelMainLeft.Controls.Add(this.labelBeginTitle);
            this.panelMainLeft.Controls.Add(this.pictureBoxIconLeft);
            this.panelMainLeft.Location = new System.Drawing.Point(20, 68);
            this.panelMainLeft.Name = "panelMainLeft";
            this.panelMainLeft.Size = new System.Drawing.Size(271, 304);
            this.panelMainLeft.TabIndex = 99;
            // 
            // labelBeginTitle
            // 
            this.labelBeginTitle.AutoSize = true;
            this.labelBeginTitle.Font = new System.Drawing.Font("Segoe UI", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelBeginTitle.ForeColor = System.Drawing.Color.White;
            this.labelBeginTitle.Location = new System.Drawing.Point(39, 9);
            this.labelBeginTitle.Name = "labelBeginTitle";
            this.labelBeginTitle.Size = new System.Drawing.Size(82, 21);
            this.labelBeginTitle.TabIndex = 79;
            this.labelBeginTitle.Text = "Welcome";
            // 
            // pictureBoxIconLeft
            // 
            this.pictureBoxIconLeft.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxIconLeft.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxIconLeft.Image")));
            this.pictureBoxIconLeft.Location = new System.Drawing.Point(17, 12);
            this.pictureBoxIconLeft.Name = "pictureBoxIconLeft";
            this.pictureBoxIconLeft.Size = new System.Drawing.Size(17, 17);
            this.pictureBoxIconLeft.TabIndex = 77;
            this.pictureBoxIconLeft.TabStop = false;
            // 
            // checkBoxStartAppWithWindows
            // 
            this.checkBoxStartAppWithWindows.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBoxStartAppWithWindows.AutoSize = true;
            this.checkBoxStartAppWithWindows.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxStartAppWithWindows.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxStartAppWithWindows.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxStartAppWithWindows.Location = new System.Drawing.Point(32, 34);
            this.checkBoxStartAppWithWindows.Name = "checkBoxStartAppWithWindows";
            this.checkBoxStartAppWithWindows.Size = new System.Drawing.Size(135, 21);
            this.checkBoxStartAppWithWindows.TabIndex = 94;
            this.checkBoxStartAppWithWindows.Text = "Start with Windows";
            this.checkBoxStartAppWithWindows.UseVisualStyleBackColor = true;
            // 
            // checkBoxStartAppMinimized
            // 
            this.checkBoxStartAppMinimized.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBoxStartAppMinimized.AutoSize = true;
            this.checkBoxStartAppMinimized.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxStartAppMinimized.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxStartAppMinimized.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxStartAppMinimized.Location = new System.Drawing.Point(32, 57);
            this.checkBoxStartAppMinimized.Name = "checkBoxStartAppMinimized";
            this.checkBoxStartAppMinimized.Size = new System.Drawing.Size(114, 21);
            this.checkBoxStartAppMinimized.TabIndex = 93;
            this.checkBoxStartAppMinimized.Text = "Start minimized";
            this.checkBoxStartAppMinimized.UseVisualStyleBackColor = true;
            // 
            // checkBoxMinimizeToTray
            // 
            this.checkBoxMinimizeToTray.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBoxMinimizeToTray.AutoSize = true;
            this.checkBoxMinimizeToTray.Checked = true;
            this.checkBoxMinimizeToTray.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxMinimizeToTray.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxMinimizeToTray.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxMinimizeToTray.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxMinimizeToTray.Location = new System.Drawing.Point(32, 11);
            this.checkBoxMinimizeToTray.Name = "checkBoxMinimizeToTray";
            this.checkBoxMinimizeToTray.Size = new System.Drawing.Size(165, 21);
            this.checkBoxMinimizeToTray.TabIndex = 92;
            this.checkBoxMinimizeToTray.Text = "Send to tray at minimize";
            this.checkBoxMinimizeToTray.UseVisualStyleBackColor = true;
            // 
            // labelLanguage
            // 
            this.labelLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.labelLanguage.AutoSize = true;
            this.labelLanguage.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelLanguage.ForeColor = System.Drawing.Color.Blue;
            this.labelLanguage.Location = new System.Drawing.Point(643, 14);
            this.labelLanguage.Name = "labelLanguage";
            this.labelLanguage.Size = new System.Drawing.Size(65, 17);
            this.labelLanguage.TabIndex = 97;
            this.labelLanguage.Text = "Language";
            // 
            // comboBoxLanguage
            // 
            this.comboBoxLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxLanguage.BackColor = System.Drawing.Color.Orange;
            this.comboBoxLanguage.DisplayMember = "DisplayName";
            this.comboBoxLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxLanguage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBoxLanguage.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxLanguage.FormattingEnabled = true;
            this.comboBoxLanguage.Location = new System.Drawing.Point(720, 12);
            this.comboBoxLanguage.Name = "comboBoxLanguage";
            this.comboBoxLanguage.Size = new System.Drawing.Size(176, 25);
            this.comboBoxLanguage.TabIndex = 96;
            this.comboBoxLanguage.SelectedIndexChanged += new System.EventHandler(this.comboBoxLanguage_SelectedIndexChanged);
            // 
            // pictureBoxLogo
            // 
            this.pictureBoxLogo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBoxLogo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBoxLogo.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxLogo.Image")));
            this.pictureBoxLogo.Location = new System.Drawing.Point(293, 48);
            this.pictureBoxLogo.Name = "pictureBoxLogo";
            this.pictureBoxLogo.Size = new System.Drawing.Size(338, 324);
            this.pictureBoxLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBoxLogo.TabIndex = 95;
            this.pictureBoxLogo.TabStop = false;
            this.pictureBoxLogo.Visible = false;
            this.pictureBoxLogo.Click += new System.EventHandler(this.pictureBoxLogo_Click);
            // 
            // panelMainRight
            // 
            this.panelMainRight.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.panelMainRight.Controls.Add(this.labelMailMe);
            this.panelMainRight.Controls.Add(this.linkLabelMailMe);
            this.panelMainRight.Controls.Add(this.labelQuestion);
            this.panelMainRight.Controls.Add(this.linkLabelNetOfficeQuestions);
            this.panelMainRight.Controls.Add(this.labelUpdate);
            this.panelMainRight.Controls.Add(this.pictureBoxIconRight);
            this.panelMainRight.Controls.Add(this.labelBug);
            this.panelMainRight.Controls.Add(this.linkLabelNetOfficeIssues);
            this.panelMainRight.Controls.Add(this.labelIWant);
            this.panelMainRight.Controls.Add(this.linkLabelNetOfficeUpdates);
            this.panelMainRight.Location = new System.Drawing.Point(633, 68);
            this.panelMainRight.Name = "panelMainRight";
            this.panelMainRight.Size = new System.Drawing.Size(271, 304);
            this.panelMainRight.TabIndex = 98;
            // 
            // labelMailMe
            // 
            this.labelMailMe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelMailMe.AutoSize = true;
            this.labelMailMe.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMailMe.ForeColor = System.Drawing.Color.Black;
            this.labelMailMe.Location = new System.Drawing.Point(12, 217);
            this.labelMailMe.Name = "labelMailMe";
            this.labelMailMe.Size = new System.Drawing.Size(118, 17);
            this.labelMailMe.TabIndex = 81;
            this.labelMailMe.Text = "make a suggestion";
            // 
            // linkLabelMailMe
            // 
            this.linkLabelMailMe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelMailMe.AutoSize = true;
            this.linkLabelMailMe.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelMailMe.Location = new System.Drawing.Point(13, 235);
            this.linkLabelMailMe.Name = "linkLabelMailMe";
            this.linkLabelMailMe.Size = new System.Drawing.Size(194, 17);
            this.linkLabelMailMe.TabIndex = 80;
            this.linkLabelMailMe.TabStop = true;
            this.linkLabelMailMe.Text = "mailto:public.sebastian@web.de";
            this.linkLabelMailMe.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel_LinkClicked);
            // 
            // labelQuestion
            // 
            this.labelQuestion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelQuestion.AutoSize = true;
            this.labelQuestion.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelQuestion.ForeColor = System.Drawing.Color.Black;
            this.labelQuestion.Location = new System.Drawing.Point(13, 158);
            this.labelQuestion.Name = "labelQuestion";
            this.labelQuestion.Size = new System.Drawing.Size(92, 17);
            this.labelQuestion.TabIndex = 79;
            this.labelQuestion.Text = "ask a question";
            // 
            // linkLabelNetOfficeQuestions
            // 
            this.linkLabelNetOfficeQuestions.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelNetOfficeQuestions.AutoSize = true;
            this.linkLabelNetOfficeQuestions.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelNetOfficeQuestions.Location = new System.Drawing.Point(14, 176);
            this.linkLabelNetOfficeQuestions.Name = "linkLabelNetOfficeQuestions";
            this.linkLabelNetOfficeQuestions.Size = new System.Drawing.Size(249, 17);
            this.linkLabelNetOfficeQuestions.TabIndex = 78;
            this.linkLabelNetOfficeQuestions.TabStop = true;
            this.linkLabelNetOfficeQuestions.Text = "http://netoffice.codeplex.com/discussions";
            this.linkLabelNetOfficeQuestions.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel_LinkClicked);
            // 
            // labelUpdate
            // 
            this.labelUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelUpdate.AutoSize = true;
            this.labelUpdate.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelUpdate.ForeColor = System.Drawing.Color.Black;
            this.labelUpdate.Location = new System.Drawing.Point(15, 101);
            this.labelUpdate.Name = "labelUpdate";
            this.labelUpdate.Size = new System.Drawing.Size(124, 17);
            this.labelUpdate.TabIndex = 75;
            this.labelUpdate.Text = "check for an update";
            // 
            // pictureBoxIconRight
            // 
            this.pictureBoxIconRight.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxIconRight.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxIconRight.Image")));
            this.pictureBoxIconRight.Location = new System.Drawing.Point(17, 12);
            this.pictureBoxIconRight.Name = "pictureBoxIconRight";
            this.pictureBoxIconRight.Size = new System.Drawing.Size(17, 17);
            this.pictureBoxIconRight.TabIndex = 77;
            this.pictureBoxIconRight.TabStop = false;
            // 
            // labelBug
            // 
            this.labelBug.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelBug.AutoSize = true;
            this.labelBug.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelBug.ForeColor = System.Drawing.Color.Black;
            this.labelBug.Location = new System.Drawing.Point(14, 48);
            this.labelBug.Name = "labelBug";
            this.labelBug.Size = new System.Drawing.Size(83, 17);
            this.labelBug.TabIndex = 74;
            this.labelBug.Text = "report a bug";
            // 
            // linkLabelNetOfficeIssues
            // 
            this.linkLabelNetOfficeIssues.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelNetOfficeIssues.AutoSize = true;
            this.linkLabelNetOfficeIssues.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelNetOfficeIssues.Location = new System.Drawing.Point(15, 66);
            this.linkLabelNetOfficeIssues.Name = "linkLabelNetOfficeIssues";
            this.linkLabelNetOfficeIssues.Size = new System.Drawing.Size(218, 17);
            this.linkLabelNetOfficeIssues.TabIndex = 10;
            this.linkLabelNetOfficeIssues.TabStop = true;
            this.linkLabelNetOfficeIssues.Text = "http://netoffice.codeplex.com/issues";
            this.linkLabelNetOfficeIssues.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel_LinkClicked);
            // 
            // labelIWant
            // 
            this.labelIWant.AutoSize = true;
            this.labelIWant.Font = new System.Drawing.Font("Segoe UI", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelIWant.ForeColor = System.Drawing.Color.White;
            this.labelIWant.Location = new System.Drawing.Point(39, 9);
            this.labelIWant.Name = "labelIWant";
            this.labelIWant.Size = new System.Drawing.Size(77, 21);
            this.labelIWant.TabIndex = 76;
            this.labelIWant.Text = "I want to";
            // 
            // linkLabelNetOfficeUpdates
            // 
            this.linkLabelNetOfficeUpdates.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelNetOfficeUpdates.AutoSize = true;
            this.linkLabelNetOfficeUpdates.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelNetOfficeUpdates.Location = new System.Drawing.Point(16, 119);
            this.linkLabelNetOfficeUpdates.Name = "linkLabelNetOfficeUpdates";
            this.linkLabelNetOfficeUpdates.Size = new System.Drawing.Size(231, 17);
            this.linkLabelNetOfficeUpdates.TabIndex = 12;
            this.linkLabelNetOfficeUpdates.TabStop = true;
            this.linkLabelNetOfficeUpdates.Text = "http://netoffice.codeplex.com/releases";
            this.linkLabelNetOfficeUpdates.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel_LinkClicked);
            // 
            // panelOptions
            // 
            this.panelOptions.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelOptions.Controls.Add(this.buttonLanguageEditor);
            this.panelOptions.Controls.Add(this.checkBoxStartAppWithWindows);
            this.panelOptions.Controls.Add(this.checkBoxStartAppMinimized);
            this.panelOptions.Controls.Add(this.checkBoxMinimizeToTray);
            this.panelOptions.Controls.Add(this.labelLanguage);
            this.panelOptions.Controls.Add(this.comboBoxLanguage);
            this.panelOptions.Location = new System.Drawing.Point(0, 402);
            this.panelOptions.Name = "panelOptions";
            this.panelOptions.Size = new System.Drawing.Size(924, 94);
            this.panelOptions.TabIndex = 102;
            // 
            // buttonLanguageEditor
            // 
            this.buttonLanguageEditor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonLanguageEditor.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonLanguageEditor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonLanguageEditor.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonLanguageEditor.ForeColor = System.Drawing.Color.Blue;
            this.buttonLanguageEditor.Image = ((System.Drawing.Image)(resources.GetObject("buttonLanguageEditor.Image")));
            this.buttonLanguageEditor.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonLanguageEditor.Location = new System.Drawing.Point(720, 53);
            this.buttonLanguageEditor.Name = "buttonLanguageEditor";
            this.buttonLanguageEditor.Size = new System.Drawing.Size(176, 25);
            this.buttonLanguageEditor.TabIndex = 98;
            this.buttonLanguageEditor.Text = "Language Editor";
            this.buttonLanguageEditor.UseVisualStyleBackColor = true;
            this.buttonLanguageEditor.Click += new System.EventHandler(this.buttonLanguageEditor_Click);
            // 
            // pictureBoxHeader
            // 
            this.pictureBoxHeader.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxHeader.Image")));
            this.pictureBoxHeader.Location = new System.Drawing.Point(185, 13);
            this.pictureBoxHeader.Name = "pictureBoxHeader";
            this.pictureBoxHeader.Size = new System.Drawing.Size(599, 36);
            this.pictureBoxHeader.TabIndex = 103;
            this.pictureBoxHeader.TabStop = false;
            this.pictureBoxHeader.Visible = false;
            // 
            // labelBeginBottom
            // 
            this.labelBeginBottom.BackColor = System.Drawing.Color.LightSteelBlue;
            this.labelBeginBottom.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelBeginBottom.Cursor = System.Windows.Forms.Cursors.Default;
            this.labelBeginBottom.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.labelBeginBottom.ForeColor = System.Drawing.Color.Black;
            this.labelBeginBottom.Location = new System.Drawing.Point(12, 147);
            this.labelBeginBottom.Name = "labelBeginBottom";
            this.labelBeginBottom.ReadOnly = true;
            this.labelBeginBottom.SelectionAlignment = NetOffice.DeveloperToolbox.Controls.Text.TextAlign.Justify;
            this.labelBeginBottom.Size = new System.Drawing.Size(255, 104);
            this.labelBeginBottom.TabIndex = 105;
            this.labelBeginBottom.Text = "You can find a help button in the upper right corner of every tab. I am looking f" +
                "orward to your message if you have any questions, suggestions, comments or reque" +
                "sts regarding the Developer Toolbox.";
            // 
            // labelBeginTop
            // 
            this.labelBeginTop.BackColor = System.Drawing.Color.LightSteelBlue;
            this.labelBeginTop.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelBeginTop.Cursor = System.Windows.Forms.Cursors.Default;
            this.labelBeginTop.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.labelBeginTop.ForeColor = System.Drawing.Color.Black;
            this.labelBeginTop.Location = new System.Drawing.Point(12, 51);
            this.labelBeginTop.Name = "labelBeginTop";
            this.labelBeginTop.ReadOnly = true;
            this.labelBeginTop.SelectionAlignment = NetOffice.DeveloperToolbox.Controls.Text.TextAlign.Justify;
            this.labelBeginTop.Size = new System.Drawing.Size(255, 93);
            this.labelBeginTop.TabIndex = 104;
            this.labelBeginTop.Text = "The NetOffice Developer Toolbox supports .NET Office developers in his daily work" +
                " with a set of helpful functions.";
            // 
            // controlForeColorAnimator1
            // 
            this.controlForeColorAnimator1.Control = this.labelVersionHint;
            this.controlForeColorAnimator1.EndColor = System.Drawing.Color.DimGray;
            this.controlForeColorAnimator1.Intervall = 5;
            this.controlForeColorAnimator1.LoopMode = NetOffice.DeveloperToolbox.Utils.Animation.LoopMode.Bidirectional;
            this.controlForeColorAnimator1.StartColor = System.Drawing.Color.Gray;
            this.controlForeColorAnimator1.StepSize = 5D;
            // 
            // WelcomeControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.pictureBoxHeader);
            this.Controls.Add(this.panelOptions);
            this.Controls.Add(this.labelVersionHint);
            this.Controls.Add(this.panelMainLeft);
            this.Controls.Add(this.pictureBoxLogo);
            this.Controls.Add(this.panelMainRight);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "WelcomeControl";
            this.Size = new System.Drawing.Size(924, 496);
            this.Resize += new System.EventHandler(this.WelcomeControl_Resize);
            this.panelMainLeft.ResumeLayout(false);
            this.panelMainLeft.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxIconLeft)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLogo)).EndInit();
            this.panelMainRight.ResumeLayout(false);
            this.panelMainRight.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxIconRight)).EndInit();
            this.panelOptions.ResumeLayout(false);
            this.panelOptions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.controlForeColorAnimator1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelVersionHint;
        private System.Windows.Forms.Panel panelMainLeft;
        private System.Windows.Forms.Label labelBeginTitle;
        private System.Windows.Forms.PictureBox pictureBoxIconLeft;
        private System.Windows.Forms.CheckBox checkBoxStartAppWithWindows;
        private System.Windows.Forms.CheckBox checkBoxStartAppMinimized;
        private System.Windows.Forms.CheckBox checkBoxMinimizeToTray;
        private System.Windows.Forms.Label labelLanguage;
        private System.Windows.Forms.ComboBox comboBoxLanguage;
        private System.Windows.Forms.PictureBox pictureBoxLogo;
        private System.Windows.Forms.Panel panelMainRight;
        private System.Windows.Forms.Label labelMailMe;
        private System.Windows.Forms.LinkLabel linkLabelMailMe;
        private System.Windows.Forms.Label labelQuestion;
        private System.Windows.Forms.LinkLabel linkLabelNetOfficeQuestions;
        private System.Windows.Forms.Label labelUpdate;
        private System.Windows.Forms.PictureBox pictureBoxIconRight;
        private System.Windows.Forms.Label labelBug;
        private System.Windows.Forms.LinkLabel linkLabelNetOfficeIssues;
        private System.Windows.Forms.Label labelIWant;
        private System.Windows.Forms.LinkLabel linkLabelNetOfficeUpdates;
        private System.Windows.Forms.Panel panelOptions;
        private System.Windows.Forms.PictureBox pictureBoxHeader;
        private NetOffice.DeveloperToolbox.Controls.Text.AdvRichTextBox labelBeginTop;
        private NetOffice.DeveloperToolbox.Controls.Text.AdvRichTextBox labelBeginBottom;
        private Utils.Animation.ControlForeColorAnimator controlForeColorAnimator1;
        private System.Windows.Forms.Button buttonLanguageEditor;
    }
}
