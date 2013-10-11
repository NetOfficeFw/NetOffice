namespace NetOffice.DeveloperToolbox
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
            this.radioButtonClassLibrary = new System.Windows.Forms.RadioButton();
            this.radioButtonWindowsForms = new System.Windows.Forms.RadioButton();
            this.radioButtonConsole = new System.Windows.Forms.RadioButton();
            this.labelProjectType = new System.Windows.Forms.Label();
            this.radioButtonAutomationAddin = new System.Windows.Forms.RadioButton();
            this.labelNoAdminHint = new System.Windows.Forms.Label();
            this.buttonChooseFolder = new System.Windows.Forms.Button();
            this.textBoxCustomFolder = new System.Windows.Forms.TextBox();
            this.radioButtonCustomFolder = new System.Windows.Forms.RadioButton();
            this.radioButtonVSProjectFolder = new System.Windows.Forms.RadioButton();
            this.radioButtonDesktop = new System.Windows.Forms.RadioButton();
            this.radioButtonUserFolder = new System.Windows.Forms.RadioButton();
            this.radioButtonApplicationData = new System.Windows.Forms.RadioButton();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.labelFolder = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.checkBoxUseTools = new System.Windows.Forms.CheckBox();
            this.linkLabelNSTOInfo = new System.Windows.Forms.LinkLabel();
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
            this.radioButtonClassLibrary.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonClassLibrary.Location = new System.Drawing.Point(522, 60);
            this.radioButtonClassLibrary.Name = "radioButtonClassLibrary";
            this.radioButtonClassLibrary.Size = new System.Drawing.Size(132, 20);
            this.radioButtonClassLibrary.TabIndex = 98;
            this.radioButtonClassLibrary.Text = "Klassenbibliothek";
            this.radioButtonClassLibrary.UseVisualStyleBackColor = true;
            this.radioButtonClassLibrary.CheckedChanged += new System.EventHandler(this.radioButtonProjectType_CheckedChanged);
            // 
            // radioButtonWindowsForms
            // 
            this.radioButtonWindowsForms.AutoSize = true;
            this.radioButtonWindowsForms.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonWindowsForms.Location = new System.Drawing.Point(209, 60);
            this.radioButtonWindowsForms.Name = "radioButtonWindowsForms";
            this.radioButtonWindowsForms.Size = new System.Drawing.Size(167, 20);
            this.radioButtonWindowsForms.TabIndex = 97;
            this.radioButtonWindowsForms.Text = "Windows Forms Projekt";
            this.radioButtonWindowsForms.UseVisualStyleBackColor = true;
            this.radioButtonWindowsForms.CheckedChanged += new System.EventHandler(this.radioButtonProjectType_CheckedChanged);
            // 
            // radioButtonConsole
            // 
            this.radioButtonConsole.AutoSize = true;
            this.radioButtonConsole.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonConsole.Location = new System.Drawing.Point(388, 60);
            this.radioButtonConsole.Name = "radioButtonConsole";
            this.radioButtonConsole.Size = new System.Drawing.Size(127, 20);
            this.radioButtonConsole.TabIndex = 96;
            this.radioButtonConsole.Text = "Konsolen Projekt";
            this.radioButtonConsole.UseVisualStyleBackColor = true;
            this.radioButtonConsole.CheckedChanged += new System.EventHandler(this.radioButtonProjectType_CheckedChanged);
            // 
            // labelProjectType
            // 
            this.labelProjectType.AutoSize = true;
            this.labelProjectType.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.labelProjectType.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProjectType.Location = new System.Drawing.Point(67, 33);
            this.labelProjectType.Name = "labelProjectType";
            this.labelProjectType.Size = new System.Drawing.Size(113, 16);
            this.labelProjectType.TabIndex = 95;
            this.labelProjectType.Text = "Typ des Projekts:";
            // 
            // radioButtonAutomationAddin
            // 
            this.radioButtonAutomationAddin.AutoSize = true;
            this.radioButtonAutomationAddin.Checked = true;
            this.radioButtonAutomationAddin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonAutomationAddin.Location = new System.Drawing.Point(73, 59);
            this.radioButtonAutomationAddin.Name = "radioButtonAutomationAddin";
            this.radioButtonAutomationAddin.Size = new System.Drawing.Size(131, 20);
            this.radioButtonAutomationAddin.TabIndex = 94;
            this.radioButtonAutomationAddin.TabStop = true;
            this.radioButtonAutomationAddin.Text = "Automation Addin";
            this.radioButtonAutomationAddin.UseVisualStyleBackColor = true;
            this.radioButtonAutomationAddin.CheckedChanged += new System.EventHandler(this.radioButtonProjectType_CheckedChanged);
            // 
            // labelNoAdminHint
            // 
            this.labelNoAdminHint.AutoSize = true;
            this.labelNoAdminHint.BackColor = System.Drawing.Color.DarkKhaki;
            this.labelNoAdminHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelNoAdminHint.Location = new System.Drawing.Point(40, 148);
            this.labelNoAdminHint.Name = "labelNoAdminHint";
            this.labelNoAdminHint.Size = new System.Drawing.Size(517, 16);
            this.labelNoAdminHint.TabIndex = 112;
            this.labelNoAdminHint.Text = "Developer Toolbox hat festgestellt das nicht für alle Ordner Schreibzugriff verfü" +
                "gbar ist.";
            this.labelNoAdminHint.Visible = false;
            // 
            // buttonChooseFolder
            // 
            this.buttonChooseFolder.Enabled = false;
            this.buttonChooseFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonChooseFolder.Location = new System.Drawing.Point(515, 104);
            this.buttonChooseFolder.Name = "buttonChooseFolder";
            this.buttonChooseFolder.Size = new System.Drawing.Size(42, 20);
            this.buttonChooseFolder.TabIndex = 111;
            this.buttonChooseFolder.Text = "...";
            this.buttonChooseFolder.UseVisualStyleBackColor = true;
            this.buttonChooseFolder.Click += new System.EventHandler(this.buttonChooseFolder_Click);
            // 
            // textBoxCustomFolder
            // 
            this.textBoxCustomFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxCustomFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxCustomFolder.Location = new System.Drawing.Point(184, 104);
            this.textBoxCustomFolder.Name = "textBoxCustomFolder";
            this.textBoxCustomFolder.ReadOnly = true;
            this.textBoxCustomFolder.Size = new System.Drawing.Size(309, 22);
            this.textBoxCustomFolder.TabIndex = 110;
            // 
            // radioButtonCustomFolder
            // 
            this.radioButtonCustomFolder.AutoSize = true;
            this.radioButtonCustomFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonCustomFolder.Location = new System.Drawing.Point(40, 105);
            this.radioButtonCustomFolder.Name = "radioButtonCustomFolder";
            this.radioButtonCustomFolder.Size = new System.Drawing.Size(125, 20);
            this.radioButtonCustomFolder.TabIndex = 109;
            this.radioButtonCustomFolder.Text = "Benutzerdefiniert";
            this.radioButtonCustomFolder.UseVisualStyleBackColor = true;
            this.radioButtonCustomFolder.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // radioButtonVSProjectFolder
            // 
            this.radioButtonVSProjectFolder.AutoSize = true;
            this.radioButtonVSProjectFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonVSProjectFolder.Location = new System.Drawing.Point(184, 80);
            this.radioButtonVSProjectFolder.Name = "radioButtonVSProjectFolder";
            this.radioButtonVSProjectFolder.Size = new System.Drawing.Size(193, 20);
            this.radioButtonVSProjectFolder.TabIndex = 108;
            this.radioButtonVSProjectFolder.Text = "Visual Studio Projekt Ordner";
            this.radioButtonVSProjectFolder.UseVisualStyleBackColor = true;
            this.radioButtonVSProjectFolder.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // radioButtonDesktop
            // 
            this.radioButtonDesktop.AutoSize = true;
            this.radioButtonDesktop.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonDesktop.Location = new System.Drawing.Point(184, 54);
            this.radioButtonDesktop.Name = "radioButtonDesktop";
            this.radioButtonDesktop.Size = new System.Drawing.Size(77, 20);
            this.radioButtonDesktop.TabIndex = 107;
            this.radioButtonDesktop.Text = "Desktop";
            this.radioButtonDesktop.UseVisualStyleBackColor = true;
            this.radioButtonDesktop.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // radioButtonUserFolder
            // 
            this.radioButtonUserFolder.AutoSize = true;
            this.radioButtonUserFolder.Checked = true;
            this.radioButtonUserFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonUserFolder.Location = new System.Drawing.Point(40, 53);
            this.radioButtonUserFolder.Name = "radioButtonUserFolder";
            this.radioButtonUserFolder.Size = new System.Drawing.Size(119, 20);
            this.radioButtonUserFolder.TabIndex = 106;
            this.radioButtonUserFolder.TabStop = true;
            this.radioButtonUserFolder.Text = "Eigene Dateien";
            this.radioButtonUserFolder.UseVisualStyleBackColor = true;
            this.radioButtonUserFolder.CheckedChanged += new System.EventHandler(this.radioButtonProjectFolder_CheckedChanged);
            // 
            // radioButtonApplicationData
            // 
            this.radioButtonApplicationData.AutoSize = true;
            this.radioButtonApplicationData.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonApplicationData.Location = new System.Drawing.Point(40, 79);
            this.radioButtonApplicationData.Name = "radioButtonApplicationData";
            this.radioButtonApplicationData.Size = new System.Drawing.Size(125, 20);
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
            this.labelFolder.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.labelFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFolder.Location = new System.Drawing.Point(39, 21);
            this.labelFolder.Name = "labelFolder";
            this.labelFolder.Size = new System.Drawing.Size(189, 16);
            this.labelFolder.TabIndex = 103;
            this.labelFolder.Text = "Speicherordner für das Projekt";
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
            this.panel1.Size = new System.Drawing.Size(662, 177);
            this.panel1.TabIndex = 113;
            // 
            // checkBoxUseTools
            // 
            this.checkBoxUseTools.AutoSize = true;
            this.checkBoxUseTools.Checked = true;
            this.checkBoxUseTools.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxUseTools.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxUseTools.Location = new System.Drawing.Point(73, 95);
            this.checkBoxUseTools.Name = "checkBoxUseTools";
            this.checkBoxUseTools.Size = new System.Drawing.Size(189, 20);
            this.checkBoxUseTools.TabIndex = 114;
            this.checkBoxUseTools.Text = "NetOffice Tools verwenden";
            this.checkBoxUseTools.UseVisualStyleBackColor = true;
            // 
            // linkLabelNSTOInfo
            // 
            this.linkLabelNSTOInfo.AutoSize = true;
            this.linkLabelNSTOInfo.Location = new System.Drawing.Point(91, 125);
            this.linkLabelNSTOInfo.Name = "linkLabelNSTOInfo";
            this.linkLabelNSTOInfo.Size = new System.Drawing.Size(309, 13);
            this.linkLabelNSTOInfo.TabIndex = 115;
            this.linkLabelNSTOInfo.TabStop = true;
            this.linkLabelNSTOInfo.Text = "Gewusst wie: Verwenden der NetOffice Tools für Addin Projekte";
            this.linkLabelNSTOInfo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelNSTOInfo_LinkClicked);
            // 
            // ProjectControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.linkLabelNSTOInfo);
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
        private System.Windows.Forms.RadioButton radioButtonClassLibrary;
        private System.Windows.Forms.RadioButton radioButtonWindowsForms;
        private System.Windows.Forms.RadioButton radioButtonConsole;
        private System.Windows.Forms.Label labelProjectType;
        private System.Windows.Forms.RadioButton radioButtonAutomationAddin;
        private System.Windows.Forms.Label labelNoAdminHint;
        private System.Windows.Forms.Button buttonChooseFolder;
        private System.Windows.Forms.TextBox textBoxCustomFolder;
        private System.Windows.Forms.RadioButton radioButtonCustomFolder;
        private System.Windows.Forms.RadioButton radioButtonVSProjectFolder;
        private System.Windows.Forms.RadioButton radioButtonDesktop;
        private System.Windows.Forms.RadioButton radioButtonUserFolder;
        private System.Windows.Forms.RadioButton radioButtonApplicationData;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.Label labelFolder;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox checkBoxUseTools;
        private System.Windows.Forms.LinkLabel linkLabelNSTOInfo;
    }
}
