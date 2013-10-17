namespace NOTools.ConsoleMonitor
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
            this.TabControlMain = new System.Windows.Forms.TabControl();
            this.tabPageConsole = new System.Windows.Forms.TabPage();
            this.tabPageChannels = new System.Windows.Forms.TabPage();
            this.tabPageInfo = new System.Windows.Forms.TabPage();
            this.labelSugestions2 = new System.Windows.Forms.Label();
            this.linkLabelSuggestions1 = new System.Windows.Forms.LinkLabel();
            this.labelSugestions1 = new System.Windows.Forms.Label();
            this.pictureBoxSuggestions = new System.Windows.Forms.PictureBox();
            this.labelSuggestions = new System.Windows.Forms.Label();
            this.labelInfo5 = new System.Windows.Forms.Label();
            this.linkLabelInfo = new System.Windows.Forms.LinkLabel();
            this.labelInfo3 = new System.Windows.Forms.Label();
            this.panelSettings1 = new System.Windows.Forms.Panel();
            this.checkBoxStartInTray = new System.Windows.Forms.CheckBox();
            this.labelSettings1 = new System.Windows.Forms.Label();
            this.comboBoxAvailableScreens = new System.Windows.Forms.ComboBox();
            this.labelSettings2 = new System.Windows.Forms.Label();
            this.radioButtonMaximized = new System.Windows.Forms.RadioButton();
            this.radioButtonLastPosition = new System.Windows.Forms.RadioButton();
            this.radioButtonCenter = new System.Windows.Forms.RadioButton();
            this.radioButtonBottomRight = new System.Windows.Forms.RadioButton();
            this.radioButtonBottomLeft = new System.Windows.Forms.RadioButton();
            this.radioButtonTopRight = new System.Windows.Forms.RadioButton();
            this.radioButtonTopLeft = new System.Windows.Forms.RadioButton();
            this.pictureBoxSettings = new System.Windows.Forms.PictureBox();
            this.labelSettings = new System.Windows.Forms.Label();
            this.labelInfo2 = new System.Windows.Forms.Label();
            this.labelInfo4 = new System.Windows.Forms.Label();
            this.labelInfo = new System.Windows.Forms.Label();
            this.labelAbout = new System.Windows.Forms.Label();
            this.labelInfo1 = new System.Windows.Forms.Label();
            this.pictureBoxInfo = new System.Windows.Forms.PictureBox();
            this.labelAbout3 = new System.Windows.Forms.Label();
            this.pictureBoxAbout = new System.Windows.Forms.PictureBox();
            this.labelAbout4 = new System.Windows.Forms.Label();
            this.labelAbout2 = new System.Windows.Forms.Label();
            this.labelAbout1 = new System.Windows.Forms.Label();
            this.ImageListTabControl = new System.Windows.Forms.ImageList(this.components);
            this.StatusStripMain = new System.Windows.Forms.StatusStrip();
            this.ToolStripLabelTime = new System.Windows.Forms.ToolStripStatusLabel();
            this.StripDropDownButtonAction = new System.Windows.Forms.ToolStripDropDownButton();
            this.ToolStripMenuItemEnabled = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemAlwaysOnTop = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.ToolStripMenuItemShowAppDomain = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemShowMachine = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemShowTime = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.ToolStripMenuItemClearConsole = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemSaveContent = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemCopyContent = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripStatusLabelAttention = new System.Windows.Forms.ToolStripStatusLabel();
            this.TimerMain = new System.Windows.Forms.Timer(this.components);
            this.ContextMenuStripTray = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItemExit = new System.Windows.Forms.ToolStripMenuItem();
            this.consoleViewMain = new NOTools.ConsoleMonitor.ConsoleViewControl();
            this.channelViewMain = new NOTools.ConsoleMonitor.ChannelViewControl();
            this.TabControlMain.SuspendLayout();
            this.tabPageConsole.SuspendLayout();
            this.tabPageChannels.SuspendLayout();
            this.tabPageInfo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSuggestions)).BeginInit();
            this.panelSettings1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSettings)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxInfo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxAbout)).BeginInit();
            this.StatusStripMain.SuspendLayout();
            this.ContextMenuStripTray.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabControlMain
            // 
            this.TabControlMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.TabControlMain.Controls.Add(this.tabPageConsole);
            this.TabControlMain.Controls.Add(this.tabPageChannels);
            this.TabControlMain.Controls.Add(this.tabPageInfo);
            this.TabControlMain.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TabControlMain.ImageList = this.ImageListTabControl;
            this.TabControlMain.Location = new System.Drawing.Point(0, 0);
            this.TabControlMain.Name = "TabControlMain";
            this.TabControlMain.SelectedIndex = 0;
            this.TabControlMain.Size = new System.Drawing.Size(662, 461);
            this.TabControlMain.TabIndex = 0;
            this.TabControlMain.SelectedIndexChanged += new System.EventHandler(this.TabControlMain_SelectedIndexChanged);
            // 
            // tabPageConsole
            // 
            this.tabPageConsole.Controls.Add(this.consoleViewMain);
            this.tabPageConsole.ImageIndex = 0;
            this.tabPageConsole.Location = new System.Drawing.Point(4, 25);
            this.tabPageConsole.Name = "tabPageConsole";
            this.tabPageConsole.Size = new System.Drawing.Size(654, 432);
            this.tabPageConsole.TabIndex = 0;
            this.tabPageConsole.Text = "Console";
            // 
            // tabPageChannels
            // 
            this.tabPageChannels.Controls.Add(this.channelViewMain);
            this.tabPageChannels.ImageIndex = 1;
            this.tabPageChannels.Location = new System.Drawing.Point(4, 25);
            this.tabPageChannels.Name = "tabPageChannels";
            this.tabPageChannels.Size = new System.Drawing.Size(654, 432);
            this.tabPageChannels.TabIndex = 1;
            this.tabPageChannels.Text = "Channels";
            this.tabPageChannels.UseVisualStyleBackColor = true;
            // 
            // tabPageInfo
            // 
            this.tabPageInfo.Controls.Add(this.labelSugestions2);
            this.tabPageInfo.Controls.Add(this.linkLabelSuggestions1);
            this.tabPageInfo.Controls.Add(this.labelSugestions1);
            this.tabPageInfo.Controls.Add(this.pictureBoxSuggestions);
            this.tabPageInfo.Controls.Add(this.labelSuggestions);
            this.tabPageInfo.Controls.Add(this.labelInfo5);
            this.tabPageInfo.Controls.Add(this.linkLabelInfo);
            this.tabPageInfo.Controls.Add(this.labelInfo3);
            this.tabPageInfo.Controls.Add(this.panelSettings1);
            this.tabPageInfo.Controls.Add(this.pictureBoxSettings);
            this.tabPageInfo.Controls.Add(this.labelSettings);
            this.tabPageInfo.Controls.Add(this.labelInfo2);
            this.tabPageInfo.Controls.Add(this.labelInfo4);
            this.tabPageInfo.Controls.Add(this.labelInfo);
            this.tabPageInfo.Controls.Add(this.labelAbout);
            this.tabPageInfo.Controls.Add(this.labelInfo1);
            this.tabPageInfo.Controls.Add(this.pictureBoxInfo);
            this.tabPageInfo.Controls.Add(this.labelAbout3);
            this.tabPageInfo.Controls.Add(this.pictureBoxAbout);
            this.tabPageInfo.Controls.Add(this.labelAbout4);
            this.tabPageInfo.Controls.Add(this.labelAbout2);
            this.tabPageInfo.Controls.Add(this.labelAbout1);
            this.tabPageInfo.ImageIndex = 2;
            this.tabPageInfo.Location = new System.Drawing.Point(4, 25);
            this.tabPageInfo.Name = "tabPageInfo";
            this.tabPageInfo.Size = new System.Drawing.Size(654, 432);
            this.tabPageInfo.TabIndex = 2;
            this.tabPageInfo.Text = "Info && Settings";
            this.tabPageInfo.UseVisualStyleBackColor = true;
            // 
            // labelSugestions2
            // 
            this.labelSugestions2.AutoSize = true;
            this.labelSugestions2.ForeColor = System.Drawing.Color.Black;
            this.labelSugestions2.Location = new System.Drawing.Point(61, 393);
            this.labelSugestions2.Name = "labelSugestions2";
            this.labelSugestions2.Size = new System.Drawing.Size(231, 16);
            this.labelSugestions2.TabIndex = 27;
            this.labelSugestions2.Text = "please use the NO Discussion Board:";
            // 
            // linkLabelSuggestions1
            // 
            this.linkLabelSuggestions1.AutoSize = true;
            this.linkLabelSuggestions1.Location = new System.Drawing.Point(293, 392);
            this.linkLabelSuggestions1.Name = "linkLabelSuggestions1";
            this.linkLabelSuggestions1.Size = new System.Drawing.Size(253, 16);
            this.linkLabelSuggestions1.TabIndex = 26;
            this.linkLabelSuggestions1.TabStop = true;
            this.linkLabelSuggestions1.Text = "http://netoffice.codeplex.com/discussions";
            this.linkLabelSuggestions1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelSuggestions1_LinkClicked);
            // 
            // labelSugestions1
            // 
            this.labelSugestions1.AutoSize = true;
            this.labelSugestions1.ForeColor = System.Drawing.Color.Black;
            this.labelSugestions1.Location = new System.Drawing.Point(61, 372);
            this.labelSugestions1.Name = "labelSugestions1";
            this.labelSugestions1.Size = new System.Drawing.Size(455, 16);
            this.labelSugestions1.TabIndex = 22;
            this.labelSugestions1.Text = "NetOffice is a community driven project. If you have an issue or suggestions,";
            // 
            // pictureBoxSuggestions
            // 
            this.pictureBoxSuggestions.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxSuggestions.Image")));
            this.pictureBoxSuggestions.Location = new System.Drawing.Point(33, 348);
            this.pictureBoxSuggestions.Name = "pictureBoxSuggestions";
            this.pictureBoxSuggestions.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxSuggestions.TabIndex = 25;
            this.pictureBoxSuggestions.TabStop = false;
            // 
            // labelSuggestions
            // 
            this.labelSuggestions.AutoSize = true;
            this.labelSuggestions.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSuggestions.ForeColor = System.Drawing.Color.Black;
            this.labelSuggestions.Location = new System.Drawing.Point(61, 348);
            this.labelSuggestions.Name = "labelSuggestions";
            this.labelSuggestions.Size = new System.Drawing.Size(76, 13);
            this.labelSuggestions.TabIndex = 24;
            this.labelSuggestions.Text = "Suggestions";
            // 
            // labelInfo5
            // 
            this.labelInfo5.AutoSize = true;
            this.labelInfo5.ForeColor = System.Drawing.Color.Black;
            this.labelInfo5.Location = new System.Drawing.Point(58, 209);
            this.labelInfo5.Name = "labelInfo5";
            this.labelInfo5.Size = new System.Drawing.Size(79, 16);
            this.labelInfo5.TabIndex = 23;
            this.labelInfo5.Text = "more about:";
            // 
            // linkLabelInfo
            // 
            this.linkLabelInfo.AutoSize = true;
            this.linkLabelInfo.Location = new System.Drawing.Point(139, 209);
            this.linkLabelInfo.Name = "linkLabelInfo";
            this.linkLabelInfo.Size = new System.Drawing.Size(276, 16);
            this.linkLabelInfo.TabIndex = 22;
            this.linkLabelInfo.TabStop = true;
            this.linkLabelInfo.Tag = "http://netoffice.codeplex.com/wikipage?title=ConsoleMonitor";
            this.linkLabelInfo.Text = "http://netoffice.codeplex.com/ConsoleMonitor";
            this.linkLabelInfo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelInfo_LinkClicked);
            // 
            // labelInfo3
            // 
            this.labelInfo3.AutoSize = true;
            this.labelInfo3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelInfo3.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelInfo3.Location = new System.Drawing.Point(58, 165);
            this.labelInfo3.Name = "labelInfo3";
            this.labelInfo3.Size = new System.Drawing.Size(569, 13);
            this.labelInfo3.TabIndex = 21;
            this.labelInfo3.Text = "string NetOffice.DebugConsole.Default.SendPipeConsoleMessage(string console, stri" +
                "ng message, string parentEntryID);";
            // 
            // panelSettings1
            // 
            this.panelSettings1.Controls.Add(this.checkBoxStartInTray);
            this.panelSettings1.Controls.Add(this.labelSettings1);
            this.panelSettings1.Controls.Add(this.comboBoxAvailableScreens);
            this.panelSettings1.Controls.Add(this.labelSettings2);
            this.panelSettings1.Controls.Add(this.radioButtonMaximized);
            this.panelSettings1.Controls.Add(this.radioButtonLastPosition);
            this.panelSettings1.Controls.Add(this.radioButtonCenter);
            this.panelSettings1.Controls.Add(this.radioButtonBottomRight);
            this.panelSettings1.Controls.Add(this.radioButtonBottomLeft);
            this.panelSettings1.Controls.Add(this.radioButtonTopRight);
            this.panelSettings1.Controls.Add(this.radioButtonTopLeft);
            this.panelSettings1.Location = new System.Drawing.Point(58, 264);
            this.panelSettings1.Name = "panelSettings1";
            this.panelSettings1.Size = new System.Drawing.Size(564, 71);
            this.panelSettings1.TabIndex = 19;
            // 
            // checkBoxStartInTray
            // 
            this.checkBoxStartInTray.AutoSize = true;
            this.checkBoxStartInTray.Location = new System.Drawing.Point(301, 47);
            this.checkBoxStartInTray.Name = "checkBoxStartInTray";
            this.checkBoxStartInTray.Size = new System.Drawing.Size(98, 20);
            this.checkBoxStartInTray.TabIndex = 23;
            this.checkBoxStartInTray.Text = "Start in Tray";
            this.checkBoxStartInTray.UseVisualStyleBackColor = true;
            // 
            // labelSettings1
            // 
            this.labelSettings1.AutoSize = true;
            this.labelSettings1.ForeColor = System.Drawing.Color.Black;
            this.labelSettings1.Location = new System.Drawing.Point(3, 1);
            this.labelSettings1.Name = "labelSettings1";
            this.labelSettings1.Size = new System.Drawing.Size(140, 16);
            this.labelSettings1.TabIndex = 20;
            this.labelSettings1.Text = "Window Start Location";
            // 
            // comboBoxAvailableScreens
            // 
            this.comboBoxAvailableScreens.DisplayMember = "DeviceName";
            this.comboBoxAvailableScreens.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxAvailableScreens.FormattingEnabled = true;
            this.comboBoxAvailableScreens.Location = new System.Drawing.Point(399, 45);
            this.comboBoxAvailableScreens.Name = "comboBoxAvailableScreens";
            this.comboBoxAvailableScreens.Size = new System.Drawing.Size(164, 24);
            this.comboBoxAvailableScreens.TabIndex = 22;
            // 
            // labelSettings2
            // 
            this.labelSettings2.AutoSize = true;
            this.labelSettings2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSettings2.ForeColor = System.Drawing.Color.Black;
            this.labelSettings2.Location = new System.Drawing.Point(396, 25);
            this.labelSettings2.Name = "labelSettings2";
            this.labelSettings2.Size = new System.Drawing.Size(42, 13);
            this.labelSettings2.TabIndex = 21;
            this.labelSettings2.Text = "Monitor";
            // 
            // radioButtonMaximized
            // 
            this.radioButtonMaximized.AutoSize = true;
            this.radioButtonMaximized.Location = new System.Drawing.Point(300, 23);
            this.radioButtonMaximized.Name = "radioButtonMaximized";
            this.radioButtonMaximized.Size = new System.Drawing.Size(90, 20);
            this.radioButtonMaximized.TabIndex = 6;
            this.radioButtonMaximized.Text = "Maximized";
            this.radioButtonMaximized.UseVisualStyleBackColor = true;
            // 
            // radioButtonLastPosition
            // 
            this.radioButtonLastPosition.AutoSize = true;
            this.radioButtonLastPosition.Location = new System.Drawing.Point(196, 46);
            this.radioButtonLastPosition.Name = "radioButtonLastPosition";
            this.radioButtonLastPosition.Size = new System.Drawing.Size(102, 20);
            this.radioButtonLastPosition.TabIndex = 5;
            this.radioButtonLastPosition.Text = "Last Position";
            this.radioButtonLastPosition.UseVisualStyleBackColor = true;
            // 
            // radioButtonCenter
            // 
            this.radioButtonCenter.AutoSize = true;
            this.radioButtonCenter.Location = new System.Drawing.Point(196, 23);
            this.radioButtonCenter.Name = "radioButtonCenter";
            this.radioButtonCenter.Size = new System.Drawing.Size(65, 20);
            this.radioButtonCenter.TabIndex = 4;
            this.radioButtonCenter.Text = "Center";
            this.radioButtonCenter.UseVisualStyleBackColor = true;
            // 
            // radioButtonBottomRight
            // 
            this.radioButtonBottomRight.AutoSize = true;
            this.radioButtonBottomRight.Location = new System.Drawing.Point(92, 46);
            this.radioButtonBottomRight.Name = "radioButtonBottomRight";
            this.radioButtonBottomRight.Size = new System.Drawing.Size(102, 20);
            this.radioButtonBottomRight.TabIndex = 3;
            this.radioButtonBottomRight.Text = "Bottom Right";
            this.radioButtonBottomRight.UseVisualStyleBackColor = true;
            // 
            // radioButtonBottomLeft
            // 
            this.radioButtonBottomLeft.AutoSize = true;
            this.radioButtonBottomLeft.Location = new System.Drawing.Point(92, 23);
            this.radioButtonBottomLeft.Name = "radioButtonBottomLeft";
            this.radioButtonBottomLeft.Size = new System.Drawing.Size(92, 20);
            this.radioButtonBottomLeft.TabIndex = 2;
            this.radioButtonBottomLeft.Text = "Bottom Left";
            this.radioButtonBottomLeft.UseVisualStyleBackColor = true;
            // 
            // radioButtonTopRight
            // 
            this.radioButtonTopRight.AutoSize = true;
            this.radioButtonTopRight.Checked = true;
            this.radioButtonTopRight.Location = new System.Drawing.Point(6, 46);
            this.radioButtonTopRight.Name = "radioButtonTopRight";
            this.radioButtonTopRight.Size = new System.Drawing.Size(85, 20);
            this.radioButtonTopRight.TabIndex = 1;
            this.radioButtonTopRight.TabStop = true;
            this.radioButtonTopRight.Text = "Top Right";
            this.radioButtonTopRight.UseVisualStyleBackColor = true;
            // 
            // radioButtonTopLeft
            // 
            this.radioButtonTopLeft.AutoSize = true;
            this.radioButtonTopLeft.Location = new System.Drawing.Point(6, 23);
            this.radioButtonTopLeft.Name = "radioButtonTopLeft";
            this.radioButtonTopLeft.Size = new System.Drawing.Size(75, 20);
            this.radioButtonTopLeft.TabIndex = 0;
            this.radioButtonTopLeft.Text = "Top Left";
            this.radioButtonTopLeft.UseVisualStyleBackColor = true;
            // 
            // pictureBoxSettings
            // 
            this.pictureBoxSettings.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxSettings.Image")));
            this.pictureBoxSettings.Location = new System.Drawing.Point(33, 240);
            this.pictureBoxSettings.Name = "pictureBoxSettings";
            this.pictureBoxSettings.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxSettings.TabIndex = 18;
            this.pictureBoxSettings.TabStop = false;
            // 
            // labelSettings
            // 
            this.labelSettings.AutoSize = true;
            this.labelSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSettings.ForeColor = System.Drawing.Color.Black;
            this.labelSettings.Location = new System.Drawing.Point(58, 240);
            this.labelSettings.Name = "labelSettings";
            this.labelSettings.Size = new System.Drawing.Size(53, 13);
            this.labelSettings.TabIndex = 17;
            this.labelSettings.Text = "Settings";
            // 
            // labelInfo2
            // 
            this.labelInfo2.AutoSize = true;
            this.labelInfo2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelInfo2.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelInfo2.Location = new System.Drawing.Point(58, 145);
            this.labelInfo2.Name = "labelInfo2";
            this.labelInfo2.Size = new System.Drawing.Size(470, 13);
            this.labelInfo2.TabIndex = 16;
            this.labelInfo2.Text = "string NetOffice.DebugConsole.Default.SendPipeConsoleMessage(string console, stri" +
                "ng message);";
            // 
            // labelInfo4
            // 
            this.labelInfo4.AutoSize = true;
            this.labelInfo4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelInfo4.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelInfo4.Location = new System.Drawing.Point(58, 186);
            this.labelInfo4.Name = "labelInfo4";
            this.labelInfo4.Size = new System.Drawing.Size(472, 13);
            this.labelInfo4.TabIndex = 15;
            this.labelInfo4.Text = "string NetOffice.DebugConsole.Default.SendPipeChannelMessage(string channel, stri" +
                "ng message);";
            // 
            // labelInfo
            // 
            this.labelInfo.AutoSize = true;
            this.labelInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelInfo.ForeColor = System.Drawing.Color.Black;
            this.labelInfo.Location = new System.Drawing.Point(58, 98);
            this.labelInfo.Name = "labelInfo";
            this.labelInfo.Size = new System.Drawing.Size(91, 13);
            this.labelInfo.TabIndex = 14;
            this.labelInfo.Text = "Did you know?";
            // 
            // labelAbout
            // 
            this.labelAbout.AutoSize = true;
            this.labelAbout.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelAbout.ForeColor = System.Drawing.Color.Black;
            this.labelAbout.Location = new System.Drawing.Point(58, 21);
            this.labelAbout.Name = "labelAbout";
            this.labelAbout.Size = new System.Drawing.Size(40, 13);
            this.labelAbout.TabIndex = 13;
            this.labelAbout.Text = "About";
            // 
            // labelInfo1
            // 
            this.labelInfo1.AutoSize = true;
            this.labelInfo1.ForeColor = System.Drawing.Color.Black;
            this.labelInfo1.Location = new System.Drawing.Point(58, 121);
            this.labelInfo1.Name = "labelInfo1";
            this.labelInfo1.Size = new System.Drawing.Size(378, 16);
            this.labelInfo1.TabIndex = 12;
            this.labelInfo1.Text = "You can send your own notifications with the following methods.";
            // 
            // pictureBoxInfo
            // 
            this.pictureBoxInfo.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxInfo.Image")));
            this.pictureBoxInfo.Location = new System.Drawing.Point(33, 98);
            this.pictureBoxInfo.Name = "pictureBoxInfo";
            this.pictureBoxInfo.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxInfo.TabIndex = 11;
            this.pictureBoxInfo.TabStop = false;
            // 
            // labelAbout3
            // 
            this.labelAbout3.AutoSize = true;
            this.labelAbout3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelAbout3.ForeColor = System.Drawing.Color.MidnightBlue;
            this.labelAbout3.Location = new System.Drawing.Point(170, 67);
            this.labelAbout3.Name = "labelAbout3";
            this.labelAbout3.Size = new System.Drawing.Size(290, 13);
            this.labelAbout3.TabIndex = 9;
            this.labelAbout3.Text = "bool NetOffice.DebugConsole.Default.EnableSharedOutput;";
            // 
            // pictureBoxAbout
            // 
            this.pictureBoxAbout.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxAbout.Image")));
            this.pictureBoxAbout.Location = new System.Drawing.Point(33, 21);
            this.pictureBoxAbout.Name = "pictureBoxAbout";
            this.pictureBoxAbout.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxAbout.TabIndex = 8;
            this.pictureBoxAbout.TabStop = false;
            // 
            // labelAbout4
            // 
            this.labelAbout4.AutoSize = true;
            this.labelAbout4.ForeColor = System.Drawing.Color.Black;
            this.labelAbout4.Location = new System.Drawing.Point(464, 65);
            this.labelAbout4.Name = "labelAbout4";
            this.labelAbout4.Size = new System.Drawing.Size(158, 16);
            this.labelAbout4.TabIndex = 6;
            this.labelAbout4.Text = "in your application to use.";
            // 
            // labelAbout2
            // 
            this.labelAbout2.AutoSize = true;
            this.labelAbout2.ForeColor = System.Drawing.Color.Black;
            this.labelAbout2.Location = new System.Drawing.Point(58, 65);
            this.labelAbout2.Name = "labelAbout2";
            this.labelAbout2.Size = new System.Drawing.Size(112, 16);
            this.labelAbout2.TabIndex = 5;
            this.labelAbout2.Text = "Enable the option";
            // 
            // labelAbout1
            // 
            this.labelAbout1.AutoSize = true;
            this.labelAbout1.ForeColor = System.Drawing.Color.Black;
            this.labelAbout1.Location = new System.Drawing.Point(58, 43);
            this.labelAbout1.Name = "labelAbout1";
            this.labelAbout1.Size = new System.Drawing.Size(360, 16);
            this.labelAbout1.TabIndex = 4;
            this.labelAbout1.Text = "This tool is an external monitor for NetOffice.DebugConsole.";
            // 
            // ImageListTabControl
            // 
            this.ImageListTabControl.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ImageListTabControl.ImageStream")));
            this.ImageListTabControl.TransparentColor = System.Drawing.Color.Transparent;
            this.ImageListTabControl.Images.SetKeyName(0, "console.png");
            this.ImageListTabControl.Images.SetKeyName(1, "channels.png");
            this.ImageListTabControl.Images.SetKeyName(2, "information.png");
            // 
            // StatusStripMain
            // 
            this.StatusStripMain.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StatusStripMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripLabelTime,
            this.StripDropDownButtonAction,
            this.toolStripStatusLabelAttention});
            this.StatusStripMain.Location = new System.Drawing.Point(0, 461);
            this.StatusStripMain.Name = "StatusStripMain";
            this.StatusStripMain.Size = new System.Drawing.Size(662, 22);
            this.StatusStripMain.TabIndex = 1;
            this.StatusStripMain.Text = "StatusStripMain";
            // 
            // ToolStripLabelTime
            // 
            this.ToolStripLabelTime.Name = "ToolStripLabelTime";
            this.ToolStripLabelTime.Size = new System.Drawing.Size(0, 17);
            // 
            // StripDropDownButtonAction
            // 
            this.StripDropDownButtonAction.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.StripDropDownButtonAction.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripMenuItemEnabled,
            this.ToolStripMenuItemAlwaysOnTop,
            this.toolStripMenuItem1,
            this.ToolStripMenuItemShowAppDomain,
            this.ToolStripMenuItemShowMachine,
            this.ToolStripMenuItemShowTime,
            this.toolStripMenuItem3,
            this.ToolStripMenuItemClearConsole,
            this.ToolStripMenuItemSaveContent,
            this.ToolStripMenuItemCopyContent});
            this.StripDropDownButtonAction.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.StripDropDownButtonAction.Name = "StripDropDownButtonAction";
            this.StripDropDownButtonAction.Size = new System.Drawing.Size(68, 20);
            this.StripDropDownButtonAction.Text = "Action...";
            this.StripDropDownButtonAction.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.StripDropDownButtonAction.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            // 
            // ToolStripMenuItemEnabled
            // 
            this.ToolStripMenuItemEnabled.Checked = true;
            this.ToolStripMenuItemEnabled.CheckOnClick = true;
            this.ToolStripMenuItemEnabled.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ToolStripMenuItemEnabled.Name = "ToolStripMenuItemEnabled";
            this.ToolStripMenuItemEnabled.ShortcutKeys = System.Windows.Forms.Keys.F5;
            this.ToolStripMenuItemEnabled.Size = new System.Drawing.Size(320, 22);
            this.ToolStripMenuItemEnabled.Text = "Monitor Enabled";
            this.ToolStripMenuItemEnabled.CheckedChanged += new System.EventHandler(this.ToolStripMenuItemEnabled_CheckedChanged);
            // 
            // ToolStripMenuItemAlwaysOnTop
            // 
            this.ToolStripMenuItemAlwaysOnTop.Checked = true;
            this.ToolStripMenuItemAlwaysOnTop.CheckOnClick = true;
            this.ToolStripMenuItemAlwaysOnTop.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ToolStripMenuItemAlwaysOnTop.Name = "ToolStripMenuItemAlwaysOnTop";
            this.ToolStripMenuItemAlwaysOnTop.ShortcutKeys = System.Windows.Forms.Keys.F6;
            this.ToolStripMenuItemAlwaysOnTop.Size = new System.Drawing.Size(320, 22);
            this.ToolStripMenuItemAlwaysOnTop.Text = "Window Always On Top";
            this.ToolStripMenuItemAlwaysOnTop.CheckedChanged += new System.EventHandler(this.ToolStripMenuItemAlwaysOnTop_CheckedChanged);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(317, 6);
            // 
            // ToolStripMenuItemShowAppDomain
            // 
            this.ToolStripMenuItemShowAppDomain.CheckOnClick = true;
            this.ToolStripMenuItemShowAppDomain.Name = "ToolStripMenuItemShowAppDomain";
            this.ToolStripMenuItemShowAppDomain.ShortcutKeys = System.Windows.Forms.Keys.F7;
            this.ToolStripMenuItemShowAppDomain.Size = new System.Drawing.Size(320, 22);
            this.ToolStripMenuItemShowAppDomain.Text = "Show AppDomain";
            this.ToolStripMenuItemShowAppDomain.CheckedChanged += new System.EventHandler(this.ToolStripMenuItemShowAppDomain_CheckedChanged);
            // 
            // ToolStripMenuItemShowMachine
            // 
            this.ToolStripMenuItemShowMachine.CheckOnClick = true;
            this.ToolStripMenuItemShowMachine.Name = "ToolStripMenuItemShowMachine";
            this.ToolStripMenuItemShowMachine.ShortcutKeys = System.Windows.Forms.Keys.F8;
            this.ToolStripMenuItemShowMachine.Size = new System.Drawing.Size(320, 22);
            this.ToolStripMenuItemShowMachine.Text = "Show Machine";
            this.ToolStripMenuItemShowMachine.CheckedChanged += new System.EventHandler(this.ToolStripMenuItemShowMachine_CheckedChanged);
            // 
            // ToolStripMenuItemShowTime
            // 
            this.ToolStripMenuItemShowTime.Checked = true;
            this.ToolStripMenuItemShowTime.CheckOnClick = true;
            this.ToolStripMenuItemShowTime.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ToolStripMenuItemShowTime.Name = "ToolStripMenuItemShowTime";
            this.ToolStripMenuItemShowTime.ShortcutKeys = System.Windows.Forms.Keys.F9;
            this.ToolStripMenuItemShowTime.Size = new System.Drawing.Size(320, 22);
            this.ToolStripMenuItemShowTime.Text = "Show Time";
            this.ToolStripMenuItemShowTime.CheckedChanged += new System.EventHandler(this.ToolStripMenuItemShowTime_CheckedChanged);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(317, 6);
            // 
            // ToolStripMenuItemClearConsole
            // 
            this.ToolStripMenuItemClearConsole.Image = ((System.Drawing.Image)(resources.GetObject("ToolStripMenuItemClearConsole.Image")));
            this.ToolStripMenuItemClearConsole.Name = "ToolStripMenuItemClearConsole";
            this.ToolStripMenuItemClearConsole.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.X)));
            this.ToolStripMenuItemClearConsole.Size = new System.Drawing.Size(320, 22);
            this.ToolStripMenuItemClearConsole.Text = "Clear Console/Channels";
            this.ToolStripMenuItemClearConsole.Click += new System.EventHandler(this.ToolStripMenuItemClearConsole_Click);
            // 
            // ToolStripMenuItemSaveContent
            // 
            this.ToolStripMenuItemSaveContent.Image = ((System.Drawing.Image)(resources.GetObject("ToolStripMenuItemSaveContent.Image")));
            this.ToolStripMenuItemSaveContent.Name = "ToolStripMenuItemSaveContent";
            this.ToolStripMenuItemSaveContent.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.S)));
            this.ToolStripMenuItemSaveContent.Size = new System.Drawing.Size(320, 22);
            this.ToolStripMenuItemSaveContent.Text = "Save All Content To Disk(Textfile)";
            this.ToolStripMenuItemSaveContent.Click += new System.EventHandler(this.ToolStripMenuItemSaveContent_Click);
            // 
            // ToolStripMenuItemCopyContent
            // 
            this.ToolStripMenuItemCopyContent.Image = ((System.Drawing.Image)(resources.GetObject("ToolStripMenuItemCopyContent.Image")));
            this.ToolStripMenuItemCopyContent.Name = "ToolStripMenuItemCopyContent";
            this.ToolStripMenuItemCopyContent.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.C)));
            this.ToolStripMenuItemCopyContent.Size = new System.Drawing.Size(320, 22);
            this.ToolStripMenuItemCopyContent.Text = "Copy All Content To Clipboard";
            this.ToolStripMenuItemCopyContent.Click += new System.EventHandler(this.ToolStripMenuItemCopyContent_Click);
            // 
            // toolStripStatusLabelAttention
            // 
            this.toolStripStatusLabelAttention.ForeColor = System.Drawing.Color.Blue;
            this.toolStripStatusLabelAttention.Name = "toolStripStatusLabelAttention";
            this.toolStripStatusLabelAttention.Size = new System.Drawing.Size(499, 17);
            this.toolStripStatusLabelAttention.Text = "Console Monitor is running with admin permssions. click this message for further " +
                "info.";
            this.toolStripStatusLabelAttention.Click += new System.EventHandler(this.ToolStripStatusLabelAttention_Click);
            // 
            // TimerMain
            // 
            this.TimerMain.Enabled = true;
            this.TimerMain.Interval = 1000;
            this.TimerMain.Tick += new System.EventHandler(this.TimerMain_Tick);
            // 
            // ContextMenuStripTray
            // 
            this.ContextMenuStripTray.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemExit});
            this.ContextMenuStripTray.Name = "ContextMenuStripTray";
            this.ContextMenuStripTray.Size = new System.Drawing.Size(143, 26);
            // 
            // toolStripMenuItemExit
            // 
            this.toolStripMenuItemExit.Image = ((System.Drawing.Image)(resources.GetObject("toolStripMenuItemExit.Image")));
            this.toolStripMenuItemExit.Name = "toolStripMenuItemExit";
            this.toolStripMenuItemExit.Size = new System.Drawing.Size(142, 22);
            this.toolStripMenuItemExit.Text = "Exit Monitor";
            this.toolStripMenuItemExit.Click += new System.EventHandler(this.ToolStripMenuItemExit_Click);
            // 
            // consoleViewMain
            // 
            this.consoleViewMain.BackColor = System.Drawing.SystemColors.Control;
            this.consoleViewMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.consoleViewMain.Location = new System.Drawing.Point(0, 0);
            this.consoleViewMain.Name = "consoleViewMain";
            this.consoleViewMain.ShowCloseButton = false;
            this.consoleViewMain.Size = new System.Drawing.Size(654, 432);
            this.consoleViewMain.TabIndex = 0;
            this.consoleViewMain.ViewStyle = NOTools.ConsoleMonitor.ConsoleViewStyle.Plain;
            // 
            // channelViewMain
            // 
            this.channelViewMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.channelViewMain.Location = new System.Drawing.Point(0, 0);
            this.channelViewMain.Name = "channelViewMain";
            this.channelViewMain.Size = new System.Drawing.Size(654, 432);
            this.channelViewMain.TabIndex = 0;
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(662, 483);
            this.Controls.Add(this.StatusStripMain);
            this.Controls.Add(this.TabControlMain);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(670, 510);
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "DebugConsole Monitor - Enabled";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMain_FormClosing);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.Resize += new System.EventHandler(this.FormMain_Resize);
            this.TabControlMain.ResumeLayout(false);
            this.tabPageConsole.ResumeLayout(false);
            this.tabPageChannels.ResumeLayout(false);
            this.tabPageInfo.ResumeLayout(false);
            this.tabPageInfo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSuggestions)).EndInit();
            this.panelSettings1.ResumeLayout(false);
            this.panelSettings1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSettings)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxInfo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxAbout)).EndInit();
            this.StatusStripMain.ResumeLayout(false);
            this.StatusStripMain.PerformLayout();
            this.ContextMenuStripTray.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl TabControlMain;
        private System.Windows.Forms.TabPage tabPageConsole;
        private System.Windows.Forms.TabPage tabPageChannels;
        private System.Windows.Forms.TabPage tabPageInfo;
        private System.Windows.Forms.StatusStrip StatusStripMain;
        private System.Windows.Forms.ToolStripStatusLabel ToolStripLabelTime;
        private System.Windows.Forms.Timer TimerMain;
        private System.Windows.Forms.Label labelAbout1;
        private System.Windows.Forms.Label labelAbout2;
        private System.Windows.Forms.Label labelAbout4;
        private System.Windows.Forms.ToolStripDropDownButton StripDropDownButtonAction;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemEnabled;
        private System.Windows.Forms.PictureBox pictureBoxAbout;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemAlwaysOnTop;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemClearConsole;
        private System.Windows.Forms.Label labelAbout3;
        private System.Windows.Forms.PictureBox pictureBoxInfo;
        private System.Windows.Forms.Label labelInfo1;
        private System.Windows.Forms.Label labelAbout;
        private System.Windows.Forms.Label labelInfo;
        private System.Windows.Forms.Label labelInfo2;
        private System.Windows.Forms.Label labelInfo4;
        private System.Windows.Forms.Label labelSettings;
        private System.Windows.Forms.PictureBox pictureBoxSettings;
        private System.Windows.Forms.Panel panelSettings1;
        private System.Windows.Forms.RadioButton radioButtonTopRight;
        private System.Windows.Forms.RadioButton radioButtonTopLeft;
        private System.Windows.Forms.RadioButton radioButtonBottomLeft;
        private System.Windows.Forms.RadioButton radioButtonBottomRight;
        private System.Windows.Forms.RadioButton radioButtonLastPosition;
        private System.Windows.Forms.RadioButton radioButtonCenter;
        private System.Windows.Forms.Label labelSettings1;
        private System.Windows.Forms.RadioButton radioButtonMaximized;
        private System.Windows.Forms.ComboBox comboBoxAvailableScreens;
        private System.Windows.Forms.Label labelSettings2;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelAttention;
        private System.Windows.Forms.Label labelInfo3;
        private System.Windows.Forms.LinkLabel linkLabelInfo;
        private System.Windows.Forms.Label labelInfo5;
        private ChannelViewControl channelViewMain;
        private ConsoleViewControl consoleViewMain;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemSaveContent;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemCopyContent;
        private System.Windows.Forms.LinkLabel linkLabelSuggestions1;
        private System.Windows.Forms.Label labelSugestions1;
        private System.Windows.Forms.PictureBox pictureBoxSuggestions;
        private System.Windows.Forms.Label labelSuggestions;
        private System.Windows.Forms.CheckBox checkBoxStartInTray;
        private System.Windows.Forms.Label labelSugestions2;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemShowAppDomain;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemShowMachine;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemShowTime;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ImageList ImageListTabControl;
        private System.Windows.Forms.ContextMenuStrip ContextMenuStripTray;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemExit;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;

    }
}

