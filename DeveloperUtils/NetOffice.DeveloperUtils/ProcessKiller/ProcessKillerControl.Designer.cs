namespace NetOffice.DeveloperUtils.ProcessKiller
{
    partial class ProcessKillerControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProcessKillerControl));
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
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.pictureBoxRunningOff = new System.Windows.Forms.PictureBox();
            this.pictureBoxRunningOn = new System.Windows.Forms.PictureBox();
            this.buttonKillApps = new System.Windows.Forms.Button();
            this.checkBoxAppKill = new System.Windows.Forms.CheckBox();
            this.listViewApps = new System.Windows.Forms.ListView();
            this.columnName = new System.Windows.Forms.ColumnHeader();
            this.columnInstances = new System.Windows.Forms.ColumnHeader();
            this.checkBoxAppsTray = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxHotKey = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonInfo = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxRunningOff)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxRunningOn)).BeginInit();
            this.SuspendLayout();
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label5.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label5.Location = new System.Drawing.Point(247, 246);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 13);
            this.label5.TabIndex = 25;
            this.label5.Text = "not running icon";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label4.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label4.Location = new System.Drawing.Point(149, 246);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 24;
            this.label4.Text = "running icon";
            // 
            // pictureBoxRunningOff
            // 
            this.pictureBoxRunningOff.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxRunningOff.Image")));
            this.pictureBoxRunningOff.Location = new System.Drawing.Point(226, 243);
            this.pictureBoxRunningOff.Name = "pictureBoxRunningOff";
            this.pictureBoxRunningOff.Size = new System.Drawing.Size(18, 18);
            this.pictureBoxRunningOff.TabIndex = 23;
            this.pictureBoxRunningOff.TabStop = false;
            // 
            // pictureBoxRunningOn
            // 
            this.pictureBoxRunningOn.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxRunningOn.Image")));
            this.pictureBoxRunningOn.Location = new System.Drawing.Point(128, 243);
            this.pictureBoxRunningOn.Name = "pictureBoxRunningOn";
            this.pictureBoxRunningOn.Size = new System.Drawing.Size(18, 18);
            this.pictureBoxRunningOn.TabIndex = 22;
            this.pictureBoxRunningOn.TabStop = false;
            // 
            // buttonKillApps
            // 
            this.buttonKillApps.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonKillApps.Image = ((System.Drawing.Image)(resources.GetObject("buttonKillApps.Image")));
            this.buttonKillApps.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonKillApps.Location = new System.Drawing.Point(241, 108);
            this.buttonKillApps.Name = "buttonKillApps";
            this.buttonKillApps.Size = new System.Drawing.Size(83, 28);
            this.buttonKillApps.TabIndex = 21;
            this.buttonKillApps.Text = "Kill now ";
            this.buttonKillApps.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonKillApps.UseVisualStyleBackColor = true;
            this.buttonKillApps.Click += new System.EventHandler(this.buttonKillApps_Click);
            // 
            // checkBoxAppKill
            // 
            this.checkBoxAppKill.AutoSize = true;
            this.checkBoxAppKill.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxAppKill.Location = new System.Drawing.Point(220, 176);
            this.checkBoxAppKill.Name = "checkBoxAppKill";
            this.checkBoxAppKill.Size = new System.Drawing.Size(78, 20);
            this.checkBoxAppKill.TabIndex = 20;
            this.checkBoxAppKill.Text = "Enabled";
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
            this.listViewApps.Location = new System.Drawing.Point(42, 50);
            this.listViewApps.Name = "listViewApps";
            this.listViewApps.Size = new System.Drawing.Size(237, 86);
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
            this.checkBoxAppsTray.Location = new System.Drawing.Point(44, 241);
            this.checkBoxAppsTray.Name = "checkBoxAppsTray";
            this.checkBoxAppsTray.Size = new System.Drawing.Size(78, 20);
            this.checkBoxAppsTray.TabIndex = 18;
            this.checkBoxAppsTray.Text = "Enabled";
            this.checkBoxAppsTray.UseVisualStyleBackColor = true;
            this.checkBoxAppsTray.CheckedChanged += new System.EventHandler(this.checkBoxAppsTray_CheckedChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Khaki;
            this.label3.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label3.Location = new System.Drawing.Point(42, 223);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(293, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Show running information in system Tray (right bottom corner)";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Khaki;
            this.label2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label2.Location = new System.Drawing.Point(42, 155);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(290, 13);
            this.label2.TabIndex = 16;
            this.label2.Text = "Type hotkey combination in textbox there kill instances";
            // 
            // textBoxHotKey
            // 
            this.textBoxHotKey.BackColor = System.Drawing.SystemColors.Control;
            this.textBoxHotKey.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxHotKey.Location = new System.Drawing.Point(42, 174);
            this.textBoxHotKey.Name = "textBoxHotKey";
            this.textBoxHotKey.ReadOnly = true;
            this.textBoxHotKey.Size = new System.Drawing.Size(153, 22);
            this.textBoxHotKey.TabIndex = 15;
            this.textBoxHotKey.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxHotKey_KeyDown);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Khaki;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label1.Location = new System.Drawing.Point(42, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(293, 13);
            this.label1.TabIndex = 14;
            this.label1.Text = "Choose the office applications you want to kill";
            // 
            // buttonInfo
            // 
            this.buttonInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonInfo.Image = ((System.Drawing.Image)(resources.GetObject("buttonInfo.Image")));
            this.buttonInfo.Location = new System.Drawing.Point(440, 5);
            this.buttonInfo.Name = "buttonInfo";
            this.buttonInfo.Size = new System.Drawing.Size(28, 28);
            this.buttonInfo.TabIndex = 26;
            this.buttonInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonInfo.UseVisualStyleBackColor = true;
            this.buttonInfo.Click += new System.EventHandler(this.buttonInfo_Click);
            // 
            // ProcessKillerControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.buttonInfo);
            this.Controls.Add(this.buttonKillApps);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pictureBoxRunningOff);
            this.Controls.Add(this.pictureBoxRunningOn);
            this.Controls.Add(this.checkBoxAppKill);
            this.Controls.Add(this.listViewApps);
            this.Controls.Add(this.checkBoxAppsTray);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxHotKey);
            this.Controls.Add(this.label1);
            this.Name = "ProcessKillerControl";
            this.Size = new System.Drawing.Size(476, 311);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxRunningOff)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxRunningOn)).EndInit();
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

        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.PictureBox pictureBoxRunningOff;
        private System.Windows.Forms.PictureBox pictureBoxRunningOn;
        private System.Windows.Forms.Button buttonKillApps;
        private System.Windows.Forms.CheckBox checkBoxAppKill;
        private System.Windows.Forms.ListView listViewApps;
        private System.Windows.Forms.ColumnHeader columnName;
        private System.Windows.Forms.ColumnHeader columnInstances;
        private System.Windows.Forms.CheckBox checkBoxAppsTray;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxHotKey;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonInfo;
    }
}
