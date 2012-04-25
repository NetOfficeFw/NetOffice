namespace NetOffice.DeveloperToolbox.ApplicationObserver
{
    partial class ApplicationObserverControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
 
        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ApplicationObserverControl));
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "Excel",
            ""}, -1);
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem(new string[] {
            "Winword",
            ""}, -1);
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem(new string[] {
            "Outlook",
            ""}, -1);
            System.Windows.Forms.ListViewItem listViewItem4 = new System.Windows.Forms.ListViewItem(new string[] {
            "PowerPnt",
            ""}, -1);
            System.Windows.Forms.ListViewItem listViewItem5 = new System.Windows.Forms.ListViewItem(new string[] {
            "MsAccess",
            ""}, -1);
            this.labelNoOfficeAppRunning = new System.Windows.Forms.Label();
            this.labelOneOrMoreIsRunning = new System.Windows.Forms.Label();
            this.pictureBoxRunningOff = new System.Windows.Forms.PictureBox();
            this.pictureBoxRunningOn = new System.Windows.Forms.PictureBox();
            this.buttonKillApps = new System.Windows.Forms.Button();
            this.checkBoxAppKill = new System.Windows.Forms.CheckBox();
            this.listViewApps = new System.Windows.Forms.ListView();
            this.columnName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnInstances = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.checkBoxAppsTray = new System.Windows.Forms.CheckBox();
            this.labelShowTray = new System.Windows.Forms.Label();
            this.labelInsertHotkey = new System.Windows.Forms.Label();
            this.textBoxHotKey = new System.Windows.Forms.TextBox();
            this.labelMain = new System.Windows.Forms.Label();
            this.buttonInfo = new System.Windows.Forms.Button();
            this.labelOfficeApplication = new System.Windows.Forms.Label();
            this.labelOfficeApplicationInstanceCount = new System.Windows.Forms.Label();
            this.labelActiveProcessList = new System.Windows.Forms.Label();
            this.checkBoxShowQuestion = new System.Windows.Forms.CheckBox();
            this.labelKillQuestion = new System.Windows.Forms.Label();
            this.listViewProcess = new System.Windows.Forms.ListView();
            this.colImage = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.pictureBox8 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxRunningOff)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxRunningOn)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox8)).BeginInit();
            this.SuspendLayout();
            // 
            // labelNoOfficeAppRunning
            // 
            this.labelNoOfficeAppRunning.AutoSize = true;
            this.labelNoOfficeAppRunning.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labelNoOfficeAppRunning.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelNoOfficeAppRunning.Location = new System.Drawing.Point(62, 389);
            this.labelNoOfficeAppRunning.Name = "labelNoOfficeAppRunning";
            this.labelNoOfficeAppRunning.Size = new System.Drawing.Size(219, 13);
            this.labelNoOfficeAppRunning.TabIndex = 25;
            this.labelNoOfficeAppRunning.Text = "Keine der ausgewählten Anwendungen aktiv";
            // 
            // labelOneOrMoreIsRunning
            // 
            this.labelOneOrMoreIsRunning.AutoSize = true;
            this.labelOneOrMoreIsRunning.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labelOneOrMoreIsRunning.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelOneOrMoreIsRunning.Location = new System.Drawing.Point(62, 366);
            this.labelOneOrMoreIsRunning.Name = "labelOneOrMoreIsRunning";
            this.labelOneOrMoreIsRunning.Size = new System.Drawing.Size(278, 13);
            this.labelOneOrMoreIsRunning.TabIndex = 24;
            this.labelOneOrMoreIsRunning.Text = "Eine oder mehrere der ausgewählten Anwendungen aktiv";
            // 
            // pictureBoxRunningOff
            // 
            this.pictureBoxRunningOff.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxRunningOff.Image")));
            this.pictureBoxRunningOff.Location = new System.Drawing.Point(41, 386);
            this.pictureBoxRunningOff.Name = "pictureBoxRunningOff";
            this.pictureBoxRunningOff.Size = new System.Drawing.Size(18, 18);
            this.pictureBoxRunningOff.TabIndex = 23;
            this.pictureBoxRunningOff.TabStop = false;
            // 
            // pictureBoxRunningOn
            // 
            this.pictureBoxRunningOn.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxRunningOn.Image")));
            this.pictureBoxRunningOn.Location = new System.Drawing.Point(41, 363);
            this.pictureBoxRunningOn.Name = "pictureBoxRunningOn";
            this.pictureBoxRunningOn.Size = new System.Drawing.Size(18, 18);
            this.pictureBoxRunningOn.TabIndex = 22;
            this.pictureBoxRunningOn.TabStop = false;
            // 
            // buttonKillApps
            // 
            this.buttonKillApps.Enabled = false;
            this.buttonKillApps.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonKillApps.Image = ((System.Drawing.Image)(resources.GetObject("buttonKillApps.Image")));
            this.buttonKillApps.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonKillApps.Location = new System.Drawing.Point(272, 164);
            this.buttonKillApps.Name = "buttonKillApps";
            this.buttonKillApps.Size = new System.Drawing.Size(119, 28);
            this.buttonKillApps.TabIndex = 21;
            this.buttonKillApps.Text = "   Beenden";
            this.buttonKillApps.UseVisualStyleBackColor = true;
            this.buttonKillApps.Click += new System.EventHandler(this.buttonKillApps_Click);
            // 
            // checkBoxAppKill
            // 
            this.checkBoxAppKill.AutoSize = true;
            this.checkBoxAppKill.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxAppKill.Location = new System.Drawing.Point(220, 255);
            this.checkBoxAppKill.Name = "checkBoxAppKill";
            this.checkBoxAppKill.Size = new System.Drawing.Size(108, 20);
            this.checkBoxAppKill.TabIndex = 20;
            this.checkBoxAppKill.Text = "Eingeschaltet";
            this.checkBoxAppKill.UseVisualStyleBackColor = true;
            this.checkBoxAppKill.CheckedChanged += new System.EventHandler(this.checkBoxAppKill_CheckedChanged);
            // 
            // listViewApps
            // 
            this.listViewApps.BackColor = System.Drawing.SystemColors.Control;
            this.listViewApps.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listViewApps.CheckBoxes = true;
            this.listViewApps.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnName,
            this.columnInstances});
            this.listViewApps.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listViewApps.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            listViewItem1.StateImageIndex = 0;
            listViewItem2.StateImageIndex = 0;
            listViewItem3.StateImageIndex = 0;
            listViewItem4.StateImageIndex = 0;
            listViewItem5.StateImageIndex = 0;
            this.listViewApps.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3,
            listViewItem4,
            listViewItem5});
            this.listViewApps.Location = new System.Drawing.Point(42, 106);
            this.listViewApps.Name = "listViewApps";
            this.listViewApps.Size = new System.Drawing.Size(224, 112);
            this.listViewApps.TabIndex = 19;
            this.listViewApps.UseCompatibleStateImageBehavior = false;
            this.listViewApps.View = System.Windows.Forms.View.Details;
            this.listViewApps.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.listViewApps_ItemChecked);
            // 
            // columnName
            // 
            this.columnName.Text = "Name";
            this.columnName.Width = 100;
            // 
            // columnInstances
            // 
            this.columnInstances.Text = "Instances";
            this.columnInstances.Width = 100;
            // 
            // checkBoxAppsTray
            // 
            this.checkBoxAppsTray.AutoSize = true;
            this.checkBoxAppsTray.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxAppsTray.Location = new System.Drawing.Point(219, 334);
            this.checkBoxAppsTray.Name = "checkBoxAppsTray";
            this.checkBoxAppsTray.Size = new System.Drawing.Size(108, 20);
            this.checkBoxAppsTray.TabIndex = 18;
            this.checkBoxAppsTray.Text = "Eingeschaltet";
            this.checkBoxAppsTray.UseVisualStyleBackColor = true;
            this.checkBoxAppsTray.CheckedChanged += new System.EventHandler(this.checkBoxAppsTray_CheckedChanged);
            // 
            // labelShowTray
            // 
            this.labelShowTray.AutoSize = true;
            this.labelShowTray.BackColor = System.Drawing.Color.Khaki;
            this.labelShowTray.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelShowTray.Location = new System.Drawing.Point(42, 338);
            this.labelShowTray.Name = "labelShowTray";
            this.labelShowTray.Size = new System.Drawing.Size(169, 13);
            this.labelShowTray.TabIndex = 17;
            this.labelShowTray.Text = "Information als Tray Icon anzeigen";
            // 
            // labelInsertHotkey
            // 
            this.labelInsertHotkey.BackColor = System.Drawing.Color.Khaki;
            this.labelInsertHotkey.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelInsertHotkey.Location = new System.Drawing.Point(42, 221);
            this.labelInsertHotkey.Name = "labelInsertHotkey";
            this.labelInsertHotkey.Size = new System.Drawing.Size(349, 29);
            this.labelInsertHotkey.TabIndex = 16;
            this.labelInsertHotkey.Text = "Geben Sie eine Tastenkombination ein mit der Sie ausgewählte Office Anwendungen a" +
                "us dem Speicher entfernen möchten";
            // 
            // textBoxHotKey
            // 
            this.textBoxHotKey.BackColor = System.Drawing.SystemColors.Control;
            this.textBoxHotKey.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxHotKey.Location = new System.Drawing.Point(42, 253);
            this.textBoxHotKey.Name = "textBoxHotKey";
            this.textBoxHotKey.ReadOnly = true;
            this.textBoxHotKey.Size = new System.Drawing.Size(153, 22);
            this.textBoxHotKey.TabIndex = 15;
            this.textBoxHotKey.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxHotKey_KeyDown);
            // 
            // labelMain
            // 
            this.labelMain.BackColor = System.Drawing.Color.Khaki;
            this.labelMain.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelMain.Location = new System.Drawing.Point(60, 50);
            this.labelMain.Name = "labelMain";
            this.labelMain.Size = new System.Drawing.Size(349, 13);
            this.labelMain.TabIndex = 14;
            this.labelMain.Text = "Wählen Sie die Office Anwendungen aus die Sie überwachen möchten";
            // 
            // buttonInfo
            // 
            this.buttonInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonInfo.Image = ((System.Drawing.Image)(resources.GetObject("buttonInfo.Image")));
            this.buttonInfo.Location = new System.Drawing.Point(753, 10);
            this.buttonInfo.Name = "buttonInfo";
            this.buttonInfo.Size = new System.Drawing.Size(28, 28);
            this.buttonInfo.TabIndex = 26;
            this.buttonInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonInfo.UseVisualStyleBackColor = true;
            this.buttonInfo.Click += new System.EventHandler(this.buttonInfo_Click);
            // 
            // labelOfficeApplication
            // 
            this.labelOfficeApplication.BackColor = System.Drawing.SystemColors.Control;
            this.labelOfficeApplication.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelOfficeApplication.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelOfficeApplication.Location = new System.Drawing.Point(46, 85);
            this.labelOfficeApplication.Name = "labelOfficeApplication";
            this.labelOfficeApplication.Size = new System.Drawing.Size(73, 15);
            this.labelOfficeApplication.TabIndex = 27;
            this.labelOfficeApplication.Text = "Anwendung";
            // 
            // labelOfficeApplicationInstanceCount
            // 
            this.labelOfficeApplicationInstanceCount.BackColor = System.Drawing.SystemColors.Control;
            this.labelOfficeApplicationInstanceCount.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelOfficeApplicationInstanceCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelOfficeApplicationInstanceCount.Location = new System.Drawing.Point(146, 86);
            this.labelOfficeApplicationInstanceCount.Name = "labelOfficeApplicationInstanceCount";
            this.labelOfficeApplicationInstanceCount.Size = new System.Drawing.Size(130, 15);
            this.labelOfficeApplicationInstanceCount.TabIndex = 28;
            this.labelOfficeApplicationInstanceCount.Text = "Instanzen im Speicher";
            // 
            // labelActiveProcessList
            // 
            this.labelActiveProcessList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelActiveProcessList.BackColor = System.Drawing.Color.Khaki;
            this.labelActiveProcessList.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelActiveProcessList.Location = new System.Drawing.Point(410, 50);
            this.labelActiveProcessList.Name = "labelActiveProcessList";
            this.labelActiveProcessList.Size = new System.Drawing.Size(341, 13);
            this.labelActiveProcessList.TabIndex = 30;
            this.labelActiveProcessList.Text = "Aktive Prozesse";
            // 
            // checkBoxShowQuestion
            // 
            this.checkBoxShowQuestion.AutoSize = true;
            this.checkBoxShowQuestion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxShowQuestion.Location = new System.Drawing.Point(220, 281);
            this.checkBoxShowQuestion.Name = "checkBoxShowQuestion";
            this.checkBoxShowQuestion.Size = new System.Drawing.Size(176, 20);
            this.checkBoxShowQuestion.TabIndex = 31;
            this.checkBoxShowQuestion.Text = "Vor Beenden nachfragen";
            this.checkBoxShowQuestion.UseVisualStyleBackColor = true;
            this.checkBoxShowQuestion.CheckedChanged += new System.EventHandler(this.checkBoxShowQuestion_CheckedChanged);
            // 
            // labelKillQuestion
            // 
            this.labelKillQuestion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.labelKillQuestion.BackColor = System.Drawing.SystemColors.Control;
            this.labelKillQuestion.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelKillQuestion.Location = new System.Drawing.Point(486, 412);
            this.labelKillQuestion.Name = "labelKillQuestion";
            this.labelKillQuestion.Size = new System.Drawing.Size(176, 12);
            this.labelKillQuestion.TabIndex = 32;
            this.labelKillQuestion.Text = "Ausgewählte Instanzen löschen?";
            this.labelKillQuestion.Visible = false;
            this.labelKillQuestion.TextChanged += new System.EventHandler(this.labelKillQuestion_TextChanged);
            // 
            // listViewProcess
            // 
            this.listViewProcess.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewProcess.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colImage,
            this.colID,
            this.colName});
            this.listViewProcess.FullRowSelect = true;
            this.listViewProcess.GridLines = true;
            this.listViewProcess.LargeImageList = this.imageList1;
            this.listViewProcess.Location = new System.Drawing.Point(411, 74);
            this.listViewProcess.MultiSelect = false;
            this.listViewProcess.Name = "listViewProcess";
            this.listViewProcess.Size = new System.Drawing.Size(340, 335);
            this.listViewProcess.SmallImageList = this.imageList1;
            this.listViewProcess.TabIndex = 33;
            this.listViewProcess.UseCompatibleStateImageBehavior = false;
            this.listViewProcess.View = System.Windows.Forms.View.Details;
            // 
            // colImage
            // 
            this.colImage.Text = "";
            this.colImage.Width = 40;
            // 
            // colID
            // 
            this.colID.Text = "ID";
            this.colID.Width = 40;
            // 
            // colName
            // 
            this.colName.Text = "Name";
            this.colName.Width = 220;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "processor.png");
            this.imageList1.Images.SetKeyName(1, "Excel.ico");
            this.imageList1.Images.SetKeyName(2, "Word.ico");
            this.imageList1.Images.SetKeyName(3, "Outlook.ico");
            this.imageList1.Images.SetKeyName(4, "PowerPoint.ico");
            this.imageList1.Images.SetKeyName(5, "Access.ico");
            // 
            // pictureBox8
            // 
            this.pictureBox8.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox8.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox8.Image")));
            this.pictureBox8.Location = new System.Drawing.Point(41, 49);
            this.pictureBox8.Name = "pictureBox8";
            this.pictureBox8.Size = new System.Drawing.Size(16, 16);
            this.pictureBox8.TabIndex = 70;
            this.pictureBox8.TabStop = false;
            // 
            // ApplicationObserverControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pictureBox8);
            this.Controls.Add(this.listViewProcess);
            this.Controls.Add(this.labelKillQuestion);
            this.Controls.Add(this.checkBoxShowQuestion);
            this.Controls.Add(this.labelActiveProcessList);
            this.Controls.Add(this.labelOfficeApplicationInstanceCount);
            this.Controls.Add(this.labelOfficeApplication);
            this.Controls.Add(this.buttonInfo);
            this.Controls.Add(this.buttonKillApps);
            this.Controls.Add(this.labelNoOfficeAppRunning);
            this.Controls.Add(this.labelOneOrMoreIsRunning);
            this.Controls.Add(this.pictureBoxRunningOff);
            this.Controls.Add(this.pictureBoxRunningOn);
            this.Controls.Add(this.checkBoxAppKill);
            this.Controls.Add(this.listViewApps);
            this.Controls.Add(this.checkBoxAppsTray);
            this.Controls.Add(this.labelShowTray);
            this.Controls.Add(this.labelInsertHotkey);
            this.Controls.Add(this.textBoxHotKey);
            this.Controls.Add(this.labelMain);
            this.Name = "ApplicationObserverControl";
            this.Size = new System.Drawing.Size(800, 429);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxRunningOff)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxRunningOn)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox8)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private System.Windows.Forms.Label labelNoOfficeAppRunning;
        private System.Windows.Forms.Label labelOneOrMoreIsRunning;
        private System.Windows.Forms.PictureBox pictureBoxRunningOff;
        private System.Windows.Forms.PictureBox pictureBoxRunningOn;
        private System.Windows.Forms.Button buttonKillApps;
        private System.Windows.Forms.CheckBox checkBoxAppKill;
        private System.Windows.Forms.ListView listViewApps;
        private System.Windows.Forms.ColumnHeader columnName;
        private System.Windows.Forms.ColumnHeader columnInstances;
        private System.Windows.Forms.CheckBox checkBoxAppsTray;
        private System.Windows.Forms.Label labelShowTray;
        private System.Windows.Forms.Label labelInsertHotkey;
        private System.Windows.Forms.TextBox textBoxHotKey;
        private System.Windows.Forms.Label labelMain;
        private System.Windows.Forms.Button buttonInfo;
        private System.Windows.Forms.Label labelOfficeApplication;
        private System.Windows.Forms.Label labelOfficeApplicationInstanceCount;
        private System.Windows.Forms.Label labelActiveProcessList;
        private System.Windows.Forms.CheckBox checkBoxShowQuestion;
        private System.Windows.Forms.Label labelKillQuestion;
        private System.Windows.Forms.ListView listViewProcess;
        private System.Windows.Forms.ColumnHeader colID;
        private System.Windows.Forms.ColumnHeader colName;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ColumnHeader colImage;
        private System.Windows.Forms.PictureBox pictureBox8;
    }
}
