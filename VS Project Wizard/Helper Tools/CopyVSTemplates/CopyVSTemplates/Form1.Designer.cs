namespace CopyVSTemplates
{
    partial class Form1
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
            this.textBoxRootFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBoxVS2008 = new System.Windows.Forms.CheckBox();
            this.checkBoxVS2010 = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonPerformAction = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonChooseFolder = new System.Windows.Forms.Button();
            this.buttonOpen2008Folder = new System.Windows.Forms.Button();
            this.buttonOpen2010Folder = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.labelNetOfficeSourceFolder = new System.Windows.Forms.Label();
            this.buttonValidateSourceFolder = new System.Windows.Forms.Button();
            this.buttonChooseVSSourceFolder = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxVSSourceFolder = new System.Windows.Forms.TextBox();
            this.buttonDeleteSourceFolder = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxRootFolder
            // 
            this.textBoxRootFolder.Location = new System.Drawing.Point(92, 30);
            this.textBoxRootFolder.Name = "textBoxRootFolder";
            this.textBoxRootFolder.Size = new System.Drawing.Size(269, 20);
            this.textBoxRootFolder.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "RootFolder";
            // 
            // checkBoxVS2008
            // 
            this.checkBoxVS2008.AutoSize = true;
            this.checkBoxVS2008.Checked = true;
            this.checkBoxVS2008.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxVS2008.Location = new System.Drawing.Point(95, 70);
            this.checkBoxVS2008.Name = "checkBoxVS2008";
            this.checkBoxVS2008.Size = new System.Drawing.Size(67, 17);
            this.checkBoxVS2008.TabIndex = 2;
            this.checkBoxVS2008.Text = "VS 2008";
            this.checkBoxVS2008.UseVisualStyleBackColor = true;
            // 
            // checkBoxVS2010
            // 
            this.checkBoxVS2010.AutoSize = true;
            this.checkBoxVS2010.Checked = true;
            this.checkBoxVS2010.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxVS2010.Location = new System.Drawing.Point(181, 70);
            this.checkBoxVS2010.Name = "checkBoxVS2010";
            this.checkBoxVS2010.Size = new System.Drawing.Size(67, 17);
            this.checkBoxVS2010.TabIndex = 3;
            this.checkBoxVS2010.Text = "VS 2010";
            this.checkBoxVS2010.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Copy to:";
            // 
            // buttonPerformAction
            // 
            this.buttonPerformAction.Location = new System.Drawing.Point(273, 66);
            this.buttonPerformAction.Name = "buttonPerformAction";
            this.buttonPerformAction.Size = new System.Drawing.Size(88, 21);
            this.buttonPerformAction.TabIndex = 5;
            this.buttonPerformAction.Text = "Go for it!";
            this.buttonPerformAction.UseVisualStyleBackColor = true;
            this.buttonPerformAction.Click += new System.EventHandler(this.buttonPerformAction_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.Location = new System.Drawing.Point(30, 132);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxLog.Size = new System.Drawing.Size(388, 129);
            this.textBoxLog.TabIndex = 6;
            this.textBoxLog.WordWrap = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 116);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(25, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Log";
            // 
            // buttonChooseFolder
            // 
            this.buttonChooseFolder.Location = new System.Drawing.Point(371, 30);
            this.buttonChooseFolder.Name = "buttonChooseFolder";
            this.buttonChooseFolder.Size = new System.Drawing.Size(47, 21);
            this.buttonChooseFolder.TabIndex = 8;
            this.buttonChooseFolder.Text = "...";
            this.buttonChooseFolder.UseVisualStyleBackColor = true;
            this.buttonChooseFolder.Click += new System.EventHandler(this.buttonChooseFolder_Click);
            // 
            // buttonOpen2008Folder
            // 
            this.buttonOpen2008Folder.Location = new System.Drawing.Point(95, 93);
            this.buttonOpen2008Folder.Name = "buttonOpen2008Folder";
            this.buttonOpen2008Folder.Size = new System.Drawing.Size(75, 21);
            this.buttonOpen2008Folder.TabIndex = 9;
            this.buttonOpen2008Folder.Text = "Open Folder";
            this.buttonOpen2008Folder.UseVisualStyleBackColor = true;
            this.buttonOpen2008Folder.Click += new System.EventHandler(this.buttonOpen2008Folder_Click);
            // 
            // buttonOpen2010Folder
            // 
            this.buttonOpen2010Folder.Location = new System.Drawing.Point(181, 93);
            this.buttonOpen2010Folder.Name = "buttonOpen2010Folder";
            this.buttonOpen2010Folder.Size = new System.Drawing.Size(75, 21);
            this.buttonOpen2010Folder.TabIndex = 10;
            this.buttonOpen2010Folder.Text = "Open Folder";
            this.buttonOpen2010Folder.UseVisualStyleBackColor = true;
            this.buttonOpen2010Folder.Click += new System.EventHandler(this.buttonOpen2010Folder_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(28, 286);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(173, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Check NetOffice Assemblies Folder";
            // 
            // labelNetOfficeSourceFolder
            // 
            this.labelNetOfficeSourceFolder.AutoSize = true;
            this.labelNetOfficeSourceFolder.Location = new System.Drawing.Point(28, 308);
            this.labelNetOfficeSourceFolder.Name = "labelNetOfficeSourceFolder";
            this.labelNetOfficeSourceFolder.Size = new System.Drawing.Size(48, 13);
            this.labelNetOfficeSourceFolder.TabIndex = 12;
            this.labelNetOfficeSourceFolder.Text = "<Folder>";
            // 
            // buttonValidateSourceFolder
            // 
            this.buttonValidateSourceFolder.Location = new System.Drawing.Point(30, 334);
            this.buttonValidateSourceFolder.Name = "buttonValidateSourceFolder";
            this.buttonValidateSourceFolder.Size = new System.Drawing.Size(202, 21);
            this.buttonValidateSourceFolder.TabIndex = 13;
            this.buttonValidateSourceFolder.Text = "Validate Folder and Assemblies";
            this.buttonValidateSourceFolder.UseVisualStyleBackColor = true;
            this.buttonValidateSourceFolder.Click += new System.EventHandler(this.buttonValidateSourceFolder_Click);
            // 
            // buttonChooseVSSourceFolder
            // 
            this.buttonChooseVSSourceFolder.Location = new System.Drawing.Point(371, 370);
            this.buttonChooseVSSourceFolder.Name = "buttonChooseVSSourceFolder";
            this.buttonChooseVSSourceFolder.Size = new System.Drawing.Size(47, 21);
            this.buttonChooseVSSourceFolder.TabIndex = 16;
            this.buttonChooseVSSourceFolder.Text = "...";
            this.buttonChooseVSSourceFolder.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(27, 375);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "RootFolder";
            // 
            // textBoxVSSourceFolder
            // 
            this.textBoxVSSourceFolder.Location = new System.Drawing.Point(92, 371);
            this.textBoxVSSourceFolder.Name = "textBoxVSSourceFolder";
            this.textBoxVSSourceFolder.Size = new System.Drawing.Size(269, 20);
            this.textBoxVSSourceFolder.TabIndex = 14;
            // 
            // buttonDeleteSourceFolder
            // 
            this.buttonDeleteSourceFolder.Location = new System.Drawing.Point(238, 334);
            this.buttonDeleteSourceFolder.Name = "buttonDeleteSourceFolder";
            this.buttonDeleteSourceFolder.Size = new System.Drawing.Size(180, 21);
            this.buttonDeleteSourceFolder.TabIndex = 17;
            this.buttonDeleteSourceFolder.Text = "Delete Folder";
            this.buttonDeleteSourceFolder.UseVisualStyleBackColor = true;
            this.buttonDeleteSourceFolder.Click += new System.EventHandler(this.buttonDeleteSourceFolder_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(439, 409);
            this.Controls.Add(this.buttonDeleteSourceFolder);
            this.Controls.Add(this.buttonChooseVSSourceFolder);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBoxVSSourceFolder);
            this.Controls.Add(this.buttonValidateSourceFolder);
            this.Controls.Add(this.labelNetOfficeSourceFolder);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.buttonOpen2010Folder);
            this.Controls.Add(this.buttonOpen2008Folder);
            this.Controls.Add(this.buttonChooseFolder);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.buttonPerformAction);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.checkBoxVS2010);
            this.Controls.Add(this.checkBoxVS2008);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxRootFolder);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Zip and copy vs templates";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxRootFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBoxVS2008;
        private System.Windows.Forms.CheckBox checkBoxVS2010;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonPerformAction;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonChooseFolder;
        private System.Windows.Forms.Button buttonOpen2008Folder;
        private System.Windows.Forms.Button buttonOpen2010Folder;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label labelNetOfficeSourceFolder;
        private System.Windows.Forms.Button buttonValidateSourceFolder;
        private System.Windows.Forms.Button buttonChooseVSSourceFolder;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxVSSourceFolder;
        private System.Windows.Forms.Button buttonDeleteSourceFolder;
    }
}

