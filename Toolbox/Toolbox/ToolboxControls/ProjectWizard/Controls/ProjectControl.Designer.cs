namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    partial class ProjectControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProjectControl));
            this.pictureBox5 = new System.Windows.Forms.PictureBox();
            this.radioButtonClassLibrary = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.radioButtonWindowsForms = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.radioButtonConsole = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.labelProjectType = new System.Windows.Forms.Label();
            this.radioButtonAutomationAddin = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.labelNoAdminHint = new System.Windows.Forms.Label();
            this.buttonChooseFolder = new System.Windows.Forms.Button();
            this.textBoxCustomFolder = new System.Windows.Forms.TextBox();
            this.radioButtonCustomFolder = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.radioButtonVSProjectFolder = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.radioButtonDesktop = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.radioButtonUserFolder = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.radioButtonApplicationData = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.labelFolder = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.checkBoxUseTools = new NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox5
            // 
            this.pictureBox5.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox5.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox5.Image")));
            this.pictureBox5.Location = new System.Drawing.Point(40, 33);
            this.pictureBox5.Name = "pictureBox5";
            this.pictureBox5.Size = new System.Drawing.Size(18, 17);
            this.pictureBox5.TabIndex = 99;
            this.pictureBox5.TabStop = false;
            // 
            // radioButtonClassLibrary
            // 
            this.radioButtonClassLibrary.AutoSize = true;
            this.radioButtonClassLibrary.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonClassLibrary.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonClassLibrary.ForeColor = System.Drawing.SystemColors.ControlText;
            this.radioButtonClassLibrary.Location = new System.Drawing.Point(522, 60);
            this.radioButtonClassLibrary.Name = "radioButtonClassLibrary";
            this.radioButtonClassLibrary.Size = new System.Drawing.Size(99, 21);
            this.radioButtonClassLibrary.TabIndex = 98;
            this.radioButtonClassLibrary.Text = "Class Library";
            this.radioButtonClassLibrary.UseVisualStyleBackColor = true;
            this.radioButtonClassLibrary.CheckedChanged += new System.EventHandler(this.radioButtonProjectType_CheckedChanged);
            // 
            // radioButtonWindowsForms
            // 
            this.radioButtonWindowsForms.AutoSize = true;
            this.radioButtonWindowsForms.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonWindowsForms.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonWindowsForms.ForeColor = System.Drawing.SystemColors.ControlText;
            this.radioButtonWindowsForms.Location = new System.Drawing.Point(209, 60);
            this.radioButtonWindowsForms.Name = "radioButtonWindowsForms";
            this.radioButtonWindowsForms.Size = new System.Drawing.Size(162, 21);
            this.radioButtonWindowsForms.TabIndex = 97;
            this.radioButtonWindowsForms.Text = "Windows Forms Project";
            this.radioButtonWindowsForms.UseVisualStyleBackColor = true;
            this.radioButtonWindowsForms.CheckedChanged += new System.EventHandler(this.radioButtonProjectType_CheckedChanged);
            // 
            // radioButtonConsole
            // 
            this.radioButtonConsole.AutoSize = true;
            this.radioButtonConsole.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonConsole.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonConsole.ForeColor = System.Drawing.SystemColors.ControlText;
            this.radioButtonConsole.Location = new System.Drawing.Point(388, 60);
            this.radioButtonConsole.Name = "radioButtonConsole";
            this.radioButtonConsole.Size = new System.Drawing.Size(116, 21);
            this.radioButtonConsole.TabIndex = 96;
            this.radioButtonConsole.Text = "Console Project";
            this.radioButtonConsole.UseVisualStyleBackColor = true;
            this.radioButtonConsole.CheckedChanged += new System.EventHandler(this.radioButtonProjectType_CheckedChanged);
            // 
            // labelProjectType
            // 
            this.labelProjectType.AutoSize = true;
            this.labelProjectType.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.labelProjectType.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProjectType.ForeColor = System.Drawing.Color.Black;
            this.labelProjectType.Location = new System.Drawing.Point(67, 33);
            this.labelProjectType.Name = "labelProjectType";
            this.labelProjectType.Size = new System.Drawing.Size(85, 16);
            this.labelProjectType.TabIndex = 95;
            this.labelProjectType.Text = "Project Type";
            // 
            // radioButtonAutomationAddin
            // 
            this.radioButtonAutomationAddin.AutoSize = true;
            this.radioButtonAutomationAddin.Checked = true;
            this.radioButtonAutomationAddin.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonAutomationAddin.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonAutomationAddin.ForeColor = System.Drawing.Color.Blue;
            this.radioButtonAutomationAddin.Location = new System.Drawing.Point(73, 59);
            this.radioButtonAutomationAddin.Name = "radioButtonAutomationAddin";
            this.radioButtonAutomationAddin.Size = new System.Drawing.Size(130, 21);
            this.radioButtonAutomationAddin.TabIndex = 94;
            this.radioButtonAutomationAddin.TabStop = true;
            this.radioButtonAutomationAddin.Text = "Automation Addin";
            this.radioButtonAutomationAddin.UseVisualStyleBackColor = true;
            this.radioButtonAutomationAddin.CheckedChanged += new System.EventHandler(this.radioButtonProjectType_CheckedChanged);
            // 
            // labelNoAdminHint
            // 
            this.labelNoAdminHint.AutoSize = true;
            this.labelNoAdminHint.BackColor = System.Drawing.Color.Orange;
            this.labelNoAdminHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelNoAdminHint.Location = new System.Drawing.Point(40, 148);
            this.labelNoAdminHint.Name = "labelNoAdminHint";
            this.labelNoAdminHint.Size = new System.Drawing.Size(538, 16);
            this.labelNoAdminHint.TabIndex = 112;
            this.labelNoAdminHint.Text = "Developer Toolbox has detected the write permissions are not available for some f" +
    "olders.";
            this.labelNoAdminHint.Visible = false;
            // 
            // buttonChooseFolder
            // 
            this.buttonChooseFolder.Enabled = false;
            this.buttonChooseFolder.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonChooseFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonChooseFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonChooseFolder.Image = ((System.Drawing.Image)(resources.GetObject("buttonChooseFolder.Image")));
            this.buttonChooseFolder.Location = new System.Drawing.Point(515, 105);
            this.buttonChooseFolder.Name = "buttonChooseFolder";
            this.buttonChooseFolder.Size = new System.Drawing.Size(42, 22);
            this.buttonChooseFolder.TabIndex = 111;
            this.buttonChooseFolder.Text = "...";
            this.buttonChooseFolder.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.buttonChooseFolder.UseVisualStyleBackColor = true;
            this.buttonChooseFolder.Click += new System.EventHandler(this.buttonChooseFolder_Click);
            // 
            // textBoxCustomFolder
            // 
            this.textBoxCustomFolder.BackColor = System.Drawing.Color.LightSteelBlue;
            this.textBoxCustomFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxCustomFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxCustomFolder.Location = new System.Drawing.Point(184, 106);
            this.textBoxCustomFolder.Name = "textBoxCustomFolder";
            this.textBoxCustomFolder.ReadOnly = true;
            this.textBoxCustomFolder.Size = new System.Drawing.Size(309, 22);
            this.textBoxCustomFolder.TabIndex = 110;
            // 
            // radioButtonCustomFolder
            // 
            this.radioButtonCustomFolder.AutoSize = true;
            this.radioButtonCustomFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonCustomFolder.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonCustomFolder.ForeColor = System.Drawing.SystemColors.ControlText;
            this.radioButtonCustomFolder.Location = new System.Drawing.Point(40, 106);
            this.radioButtonCustomFolder.Name = "radioButtonCustomFolder";
            this.radioButtonCustomFolder.Size = new System.Drawing.Size(69, 21);
            this.radioButtonCustomFolder.TabIndex = 109;
            this.radioButtonCustomFolder.Text = "Custom";
            this.radioButtonCustomFolder.UseVisualStyleBackColor = true;
            this.radioButtonCustomFolder.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // radioButtonVSProjectFolder
            // 
            this.radioButtonVSProjectFolder.AutoSize = true;
            this.radioButtonVSProjectFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonVSProjectFolder.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonVSProjectFolder.ForeColor = System.Drawing.SystemColors.ControlText;
            this.radioButtonVSProjectFolder.Location = new System.Drawing.Point(184, 77);
            this.radioButtonVSProjectFolder.Name = "radioButtonVSProjectFolder";
            this.radioButtonVSProjectFolder.Size = new System.Drawing.Size(199, 21);
            this.radioButtonVSProjectFolder.TabIndex = 108;
            this.radioButtonVSProjectFolder.Text = "VS Project Folder (if available)";
            this.radioButtonVSProjectFolder.UseVisualStyleBackColor = true;
            this.radioButtonVSProjectFolder.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // radioButtonDesktop
            // 
            this.radioButtonDesktop.AutoSize = true;
            this.radioButtonDesktop.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonDesktop.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonDesktop.ForeColor = System.Drawing.SystemColors.ControlText;
            this.radioButtonDesktop.Location = new System.Drawing.Point(184, 51);
            this.radioButtonDesktop.Name = "radioButtonDesktop";
            this.radioButtonDesktop.Size = new System.Drawing.Size(73, 21);
            this.radioButtonDesktop.TabIndex = 107;
            this.radioButtonDesktop.Text = "Desktop";
            this.radioButtonDesktop.UseVisualStyleBackColor = true;
            this.radioButtonDesktop.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // radioButtonUserFolder
            // 
            this.radioButtonUserFolder.AutoSize = true;
            this.radioButtonUserFolder.Checked = true;
            this.radioButtonUserFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonUserFolder.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonUserFolder.ForeColor = System.Drawing.Color.Blue;
            this.radioButtonUserFolder.Location = new System.Drawing.Point(40, 50);
            this.radioButtonUserFolder.Name = "radioButtonUserFolder";
            this.radioButtonUserFolder.Size = new System.Drawing.Size(93, 21);
            this.radioButtonUserFolder.TabIndex = 106;
            this.radioButtonUserFolder.TabStop = true;
            this.radioButtonUserFolder.Text = "User Folder";
            this.radioButtonUserFolder.UseVisualStyleBackColor = true;
            this.radioButtonUserFolder.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // radioButtonApplicationData
            // 
            this.radioButtonApplicationData.AutoSize = true;
            this.radioButtonApplicationData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonApplicationData.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonApplicationData.ForeColor = System.Drawing.SystemColors.ControlText;
            this.radioButtonApplicationData.Location = new System.Drawing.Point(40, 76);
            this.radioButtonApplicationData.Name = "radioButtonApplicationData";
            this.radioButtonApplicationData.Size = new System.Drawing.Size(121, 21);
            this.radioButtonApplicationData.TabIndex = 105;
            this.radioButtonApplicationData.Text = "Application Data";
            this.radioButtonApplicationData.UseVisualStyleBackColor = true;
            this.radioButtonApplicationData.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // pictureBox4
            // 
            this.pictureBox4.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox4.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox4.Image")));
            this.pictureBox4.Location = new System.Drawing.Point(15, 18);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(18, 17);
            this.pictureBox4.TabIndex = 104;
            this.pictureBox4.TabStop = false;
            // 
            // labelFolder
            // 
            this.labelFolder.AutoSize = true;
            this.labelFolder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.labelFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFolder.ForeColor = System.Drawing.Color.Black;
            this.labelFolder.Location = new System.Drawing.Point(39, 21);
            this.labelFolder.Name = "labelFolder";
            this.labelFolder.Size = new System.Drawing.Size(92, 16);
            this.labelFolder.TabIndex = 103;
            this.labelFolder.Text = "Project Folder";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.labelNoAdminHint);
            this.panel1.Controls.Add(this.buttonChooseFolder);
            this.panel1.Controls.Add(this.textBoxCustomFolder);
            this.panel1.Controls.Add(this.radioButtonCustomFolder);
            this.panel1.Controls.Add(this.radioButtonVSProjectFolder);
            this.panel1.Controls.Add(this.radioButtonDesktop);
            this.panel1.Controls.Add(this.radioButtonUserFolder);
            this.panel1.Controls.Add(this.radioButtonApplicationData);
            this.panel1.Controls.Add(this.pictureBox4);
            this.panel1.Controls.Add(this.labelFolder);
            this.panel1.Location = new System.Drawing.Point(25, 157);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(629, 177);
            this.panel1.TabIndex = 113;
            // 
            // checkBoxUseTools
            // 
            this.checkBoxUseTools.AutoSize = true;
            this.checkBoxUseTools.Checked = true;
            this.checkBoxUseTools.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxUseTools.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxUseTools.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxUseTools.ForeColor = System.Drawing.SystemColors.ControlText;
            this.checkBoxUseTools.Location = new System.Drawing.Point(72, 95);
            this.checkBoxUseTools.Name = "checkBoxUseTools";
            this.checkBoxUseTools.Size = new System.Drawing.Size(140, 21);
            this.checkBoxUseTools.TabIndex = 114;
            this.checkBoxUseTools.Text = "Use NetOffice Tools";
            this.checkBoxUseTools.UseVisualStyleBackColor = true;
            // 
            // ProjectControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.checkBoxUseTools);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pictureBox5);
            this.Controls.Add(this.radioButtonClassLibrary);
            this.Controls.Add(this.radioButtonWindowsForms);
            this.Controls.Add(this.radioButtonConsole);
            this.Controls.Add(this.labelProjectType);
            this.Controls.Add(this.radioButtonAutomationAddin);
            this.Name = "ProjectControl";
            this.Size = new System.Drawing.Size(744, 342);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox5;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonClassLibrary;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonWindowsForms;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonConsole;
        private System.Windows.Forms.Label labelProjectType;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonAutomationAddin;
        private System.Windows.Forms.Label labelNoAdminHint;
        private System.Windows.Forms.Button buttonChooseFolder;
        private System.Windows.Forms.TextBox textBoxCustomFolder;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonCustomFolder;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonVSProjectFolder;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonDesktop;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonUserFolder;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonApplicationData;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.Label labelFolder;
        private System.Windows.Forms.Panel panel1;
        private NetOffice.DeveloperToolbox.Controls.Check.GlowCheckBox checkBoxUseTools;
    }
}
