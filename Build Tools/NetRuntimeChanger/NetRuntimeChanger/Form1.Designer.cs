namespace NetRuntimeChanger
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
            this.comboBoxVersion = new System.Windows.Forms.ComboBox();
            this.comboBoxLanguage = new System.Windows.Forms.ComboBox();
            this.checkBoxChangeKeyFiles = new System.Windows.Forms.CheckBox();
            this.textBoxKeyFilesRootFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBoxFolder
            // 
            this.textBoxFolder.Location = new System.Drawing.Point(12, 60);
            this.textBoxFolder.Name = "textBoxFolder";
            this.textBoxFolder.ReadOnly = true;
            this.textBoxFolder.Size = new System.Drawing.Size(474, 20);
            this.textBoxFolder.TabIndex = 0;
            // 
            // textBoxLog
            // 
            this.textBoxLog.Location = new System.Drawing.Point(12, 86);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.Size = new System.Drawing.Size(474, 136);
            this.textBoxLog.TabIndex = 1;
            // 
            // buttonChooseFolder
            // 
            this.buttonChooseFolder.Location = new System.Drawing.Point(492, 56);
            this.buttonChooseFolder.Name = "buttonChooseFolder";
            this.buttonChooseFolder.Size = new System.Drawing.Size(45, 24);
            this.buttonChooseFolder.TabIndex = 2;
            this.buttonChooseFolder.Text = "...";
            this.buttonChooseFolder.UseVisualStyleBackColor = true;
            this.buttonChooseFolder.Click += new System.EventHandler(this.buttonChooseFolder_Click);
            // 
            // buttonStart
            // 
            this.buttonStart.Location = new System.Drawing.Point(425, 228);
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.Size = new System.Drawing.Size(61, 24);
            this.buttonStart.TabIndex = 3;
            this.buttonStart.Text = "Start";
            this.buttonStart.UseVisualStyleBackColor = true;
            this.buttonStart.Click += new System.EventHandler(this.buttonStart_Click);
            // 
            // comboBoxVersion
            // 
            this.comboBoxVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxVersion.FormattingEnabled = true;
            this.comboBoxVersion.Items.AddRange(new object[] {
            ".Net 2.0",
            ".Net 3.0",
            ".Net 3.5",
            ".Net 4.0",
            ".Net 4.5"});
            this.comboBoxVersion.Location = new System.Drawing.Point(99, 230);
            this.comboBoxVersion.Name = "comboBoxVersion";
            this.comboBoxVersion.Size = new System.Drawing.Size(172, 21);
            this.comboBoxVersion.TabIndex = 4;
            // 
            // comboBoxLanguage
            // 
            this.comboBoxLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxLanguage.FormattingEnabled = true;
            this.comboBoxLanguage.Items.AddRange(new object[] {
            "C#",
            "VB.NET"});
            this.comboBoxLanguage.Location = new System.Drawing.Point(12, 230);
            this.comboBoxLanguage.Name = "comboBoxLanguage";
            this.comboBoxLanguage.Size = new System.Drawing.Size(81, 21);
            this.comboBoxLanguage.TabIndex = 5;
            // 
            // checkBoxChangeKeyFiles
            // 
            this.checkBoxChangeKeyFiles.AutoSize = true;
            this.checkBoxChangeKeyFiles.Location = new System.Drawing.Point(19, 277);
            this.checkBoxChangeKeyFiles.Name = "checkBoxChangeKeyFiles";
            this.checkBoxChangeKeyFiles.Size = new System.Drawing.Size(105, 17);
            this.checkBoxChangeKeyFiles.TabIndex = 6;
            this.checkBoxChangeKeyFiles.Text = "Change KeyFiles";
            this.checkBoxChangeKeyFiles.UseVisualStyleBackColor = true;
            // 
            // textBoxKeyFilesRootFolder
            // 
            this.textBoxKeyFilesRootFolder.Location = new System.Drawing.Point(123, 300);
            this.textBoxKeyFilesRootFolder.Name = "textBoxKeyFilesRootFolder";
            this.textBoxKeyFilesRootFolder.ReadOnly = true;
            this.textBoxKeyFilesRootFolder.Size = new System.Drawing.Size(363, 20);
            this.textBoxKeyFilesRootFolder.TabIndex = 7;
            this.textBoxKeyFilesRootFolder.Text = "C:\\NetOffice\\KeyFiles";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 303);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "KeyFiles RootFolder";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(326, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "This tool change the target .NET runtime for all projects in a solution";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(545, 341);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxKeyFilesRootFolder);
            this.Controls.Add(this.checkBoxChangeKeyFiles);
            this.Controls.Add(this.comboBoxLanguage);
            this.Controls.Add(this.comboBoxVersion);
            this.Controls.Add(this.buttonStart);
            this.Controls.Add(this.buttonChooseFolder);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.textBoxFolder);
            this.Name = "Form1";
            this.Text = ".NetRuntimerChanger";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxFolder;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.Button buttonChooseFolder;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.ComboBox comboBoxVersion;
        private System.Windows.Forms.ComboBox comboBoxLanguage;
        private System.Windows.Forms.CheckBox checkBoxChangeKeyFiles;
        private System.Windows.Forms.TextBox textBoxKeyFilesRootFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

