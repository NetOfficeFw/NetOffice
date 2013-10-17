namespace NOBuildTools.VersionUpdater
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBoxFolder = new System.Windows.Forms.TextBox();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.buttonChooseFolder = new System.Windows.Forms.Button();
            this.buttonStart = new System.Windows.Forms.Button();
            this.comboBoxToNetVersion = new System.Windows.Forms.ComboBox();
            this.checkBoxChangeKeyFiles = new System.Windows.Forms.CheckBox();
            this.textBoxKeyFilesRootFolder = new System.Windows.Forms.TextBox();
            this.labelKeyFolder = new System.Windows.Forms.Label();
            this.checkBoxChangeNetMarker = new System.Windows.Forms.CheckBox();
            this.labelFolder = new System.Windows.Forms.Label();
            this.labelNetVersion = new System.Windows.Forms.Label();
            this.buttonLoadConfig = new System.Windows.Forms.Button();
            this.buttonSaveConfig = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxFromNetVersion = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // textBoxFolder
            // 
            this.textBoxFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxFolder.Location = new System.Drawing.Point(100, 19);
            this.textBoxFolder.Name = "textBoxFolder";
            this.textBoxFolder.Size = new System.Drawing.Size(740, 20);
            this.textBoxFolder.TabIndex = 0;
            // 
            // textBoxLog
            // 
            this.textBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxLog.Location = new System.Drawing.Point(101, 155);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.Size = new System.Drawing.Size(739, 305);
            this.textBoxLog.TabIndex = 1;
            // 
            // buttonChooseFolder
            // 
            this.buttonChooseFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonChooseFolder.Location = new System.Drawing.Point(863, 16);
            this.buttonChooseFolder.Name = "buttonChooseFolder";
            this.buttonChooseFolder.Size = new System.Drawing.Size(45, 24);
            this.buttonChooseFolder.TabIndex = 2;
            this.buttonChooseFolder.Text = "...";
            this.buttonChooseFolder.UseVisualStyleBackColor = true;
            this.buttonChooseFolder.Click += new System.EventHandler(this.buttonChooseFolder_Click);
            // 
            // buttonStart
            // 
            this.buttonStart.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonStart.Location = new System.Drawing.Point(100, 125);
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.Size = new System.Drawing.Size(540, 24);
            this.buttonStart.TabIndex = 3;
            this.buttonStart.Text = "Start";
            this.buttonStart.UseVisualStyleBackColor = true;
            this.buttonStart.Click += new System.EventHandler(this.buttonStart_Click);
            // 
            // comboBoxToNetVersion
            // 
            this.comboBoxToNetVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxToNetVersion.FormattingEnabled = true;
            this.comboBoxToNetVersion.Items.AddRange(new object[] {
            ".Net 2.0",
            ".Net 3.0",
            ".Net 3.5",
            ".Net 4.0",
            ".Net 4.5"});
            this.comboBoxToNetVersion.Location = new System.Drawing.Point(476, 51);
            this.comboBoxToNetVersion.Name = "comboBoxToNetVersion";
            this.comboBoxToNetVersion.Size = new System.Drawing.Size(172, 21);
            this.comboBoxToNetVersion.TabIndex = 4;
            // 
            // checkBoxChangeKeyFiles
            // 
            this.checkBoxChangeKeyFiles.AutoSize = true;
            this.checkBoxChangeKeyFiles.Location = new System.Drawing.Point(100, 91);
            this.checkBoxChangeKeyFiles.Name = "checkBoxChangeKeyFiles";
            this.checkBoxChangeKeyFiles.Size = new System.Drawing.Size(105, 17);
            this.checkBoxChangeKeyFiles.TabIndex = 6;
            this.checkBoxChangeKeyFiles.Text = "Change KeyFiles";
            this.checkBoxChangeKeyFiles.UseVisualStyleBackColor = true;
            // 
            // textBoxKeyFilesRootFolder
            // 
            this.textBoxKeyFilesRootFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxKeyFilesRootFolder.Location = new System.Drawing.Point(330, 89);
            this.textBoxKeyFilesRootFolder.Name = "textBoxKeyFilesRootFolder";
            this.textBoxKeyFilesRootFolder.Size = new System.Drawing.Size(510, 20);
            this.textBoxKeyFilesRootFolder.TabIndex = 7;
            this.textBoxKeyFilesRootFolder.Text = "C:\\NetOffice\\KeyFiles";
            // 
            // labelKeyFolder
            // 
            this.labelKeyFolder.AutoSize = true;
            this.labelKeyFolder.Location = new System.Drawing.Point(219, 93);
            this.labelKeyFolder.Name = "labelKeyFolder";
            this.labelKeyFolder.Size = new System.Drawing.Size(101, 13);
            this.labelKeyFolder.TabIndex = 8;
            this.labelKeyFolder.Text = "KeyFiles RootFolder";
            // 
            // checkBoxChangeNetMarker
            // 
            this.checkBoxChangeNetMarker.AutoSize = true;
            this.checkBoxChangeNetMarker.Location = new System.Drawing.Point(668, 55);
            this.checkBoxChangeNetMarker.Name = "checkBoxChangeNetMarker";
            this.checkBoxChangeNetMarker.Size = new System.Drawing.Size(116, 17);
            this.checkBoxChangeNetMarker.TabIndex = 12;
            this.checkBoxChangeNetMarker.Text = "Change NetMarker";
            this.checkBoxChangeNetMarker.UseVisualStyleBackColor = true;
            // 
            // labelFolder
            // 
            this.labelFolder.AutoSize = true;
            this.labelFolder.Location = new System.Drawing.Point(21, 22);
            this.labelFolder.Name = "labelFolder";
            this.labelFolder.Size = new System.Drawing.Size(49, 13);
            this.labelFolder.TabIndex = 13;
            this.labelFolder.Text = "Directory";
            // 
            // labelNetVersion
            // 
            this.labelNetVersion.AutoSize = true;
            this.labelNetVersion.Location = new System.Drawing.Point(384, 55);
            this.labelNetVersion.Name = "labelNetVersion";
            this.labelNetVersion.Size = new System.Drawing.Size(86, 13);
            this.labelNetVersion.TabIndex = 15;
            this.labelNetVersion.Text = "To .NET Version";
            // 
            // buttonLoadConfig
            // 
            this.buttonLoadConfig.Location = new System.Drawing.Point(653, 125);
            this.buttonLoadConfig.Name = "buttonLoadConfig";
            this.buttonLoadConfig.Size = new System.Drawing.Size(90, 24);
            this.buttonLoadConfig.TabIndex = 21;
            this.buttonLoadConfig.Text = "Load Config";
            this.buttonLoadConfig.UseVisualStyleBackColor = true;
            this.buttonLoadConfig.Click += new System.EventHandler(this.buttonLoadConfig_Click);
            // 
            // buttonSaveConfig
            // 
            this.buttonSaveConfig.Location = new System.Drawing.Point(750, 125);
            this.buttonSaveConfig.Name = "buttonSaveConfig";
            this.buttonSaveConfig.Size = new System.Drawing.Size(90, 24);
            this.buttonSaveConfig.TabIndex = 22;
            this.buttonSaveConfig.Text = "Save Config";
            this.buttonSaveConfig.UseVisualStyleBackColor = true;
            this.buttonSaveConfig.Click += new System.EventHandler(this.buttonSaveConfig_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(101, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 13);
            this.label1.TabIndex = 24;
            this.label1.Text = "From NET Version";
            // 
            // comboBoxFromNetVersion
            // 
            this.comboBoxFromNetVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxFromNetVersion.FormattingEnabled = true;
            this.comboBoxFromNetVersion.Items.AddRange(new object[] {
            ".Net 2.0",
            ".Net 3.0",
            ".Net 3.5",
            ".Net 4.0",
            ".Net 4.5"});
            this.comboBoxFromNetVersion.Location = new System.Drawing.Point(200, 51);
            this.comboBoxFromNetVersion.Name = "comboBoxFromNetVersion";
            this.comboBoxFromNetVersion.Size = new System.Drawing.Size(172, 21);
            this.comboBoxFromNetVersion.TabIndex = 23;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(925, 472);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxFromNetVersion);
            this.Controls.Add(this.buttonSaveConfig);
            this.Controls.Add(this.buttonLoadConfig);
            this.Controls.Add(this.labelNetVersion);
            this.Controls.Add(this.labelFolder);
            this.Controls.Add(this.checkBoxChangeNetMarker);
            this.Controls.Add(this.labelKeyFolder);
            this.Controls.Add(this.textBoxKeyFilesRootFolder);
            this.Controls.Add(this.checkBoxChangeKeyFiles);
            this.Controls.Add(this.comboBoxToNetVersion);
            this.Controls.Add(this.buttonStart);
            this.Controls.Add(this.buttonChooseFolder);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.textBoxFolder);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NOBuildTools.VersionUpdater";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxFolder;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.Button buttonChooseFolder;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.ComboBox comboBoxToNetVersion;
        private System.Windows.Forms.CheckBox checkBoxChangeKeyFiles;
        private System.Windows.Forms.TextBox textBoxKeyFilesRootFolder;
        private System.Windows.Forms.Label labelKeyFolder;
        private System.Windows.Forms.CheckBox checkBoxChangeNetMarker;
        private System.Windows.Forms.Label labelFolder;
        private System.Windows.Forms.Label labelNetVersion;
        private System.Windows.Forms.Button buttonLoadConfig;
        private System.Windows.Forms.Button buttonSaveConfig;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxFromNetVersion;
    }
}

