namespace TutorialsBase
{
    partial class FormOptions
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormOptions));
            this.groupBoxLanguage = new System.Windows.Forms.GroupBox();
            this.radioButtonLanguage1031 = new System.Windows.Forms.RadioButton();
            this.radioButtonLanguage1033 = new System.Windows.Forms.RadioButton();
            this.buttonDone = new System.Windows.Forms.Button();
            this.groupBoxOnlineMode = new System.Windows.Forms.GroupBox();
            this.labelDocumentationHint = new System.Windows.Forms.Label();
            this.radioButtonConnect = new System.Windows.Forms.RadioButton();
            this.radioButtonShowLink = new System.Windows.Forms.RadioButton();
            this.checkBoxSaveSettings = new System.Windows.Forms.CheckBox();
            this.groupBoxLanguage.SuspendLayout();
            this.groupBoxOnlineMode.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxLanguage
            // 
            this.groupBoxLanguage.Controls.Add(this.radioButtonLanguage1031);
            this.groupBoxLanguage.Controls.Add(this.radioButtonLanguage1033);
            this.groupBoxLanguage.Location = new System.Drawing.Point(23, 26);
            this.groupBoxLanguage.Name = "groupBoxLanguage";
            this.groupBoxLanguage.Size = new System.Drawing.Size(282, 89);
            this.groupBoxLanguage.TabIndex = 3;
            this.groupBoxLanguage.TabStop = false;
            this.groupBoxLanguage.Text = "Select Language";
            // 
            // radioButtonLanguage1031
            // 
            this.radioButtonLanguage1031.AutoSize = true;
            this.radioButtonLanguage1031.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonLanguage1031.Location = new System.Drawing.Point(19, 53);
            this.radioButtonLanguage1031.Name = "radioButtonLanguage1031";
            this.radioButtonLanguage1031.Size = new System.Drawing.Size(110, 17);
            this.radioButtonLanguage1031.TabIndex = 3;
            this.radioButtonLanguage1031.Text = "German (Deutsch)";
            this.radioButtonLanguage1031.UseVisualStyleBackColor = true;
            // 
            // radioButtonLanguage1033
            // 
            this.radioButtonLanguage1033.AutoSize = true;
            this.radioButtonLanguage1033.Checked = true;
            this.radioButtonLanguage1033.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonLanguage1033.Location = new System.Drawing.Point(19, 30);
            this.radioButtonLanguage1033.Name = "radioButtonLanguage1033";
            this.radioButtonLanguage1033.Size = new System.Drawing.Size(82, 17);
            this.radioButtonLanguage1033.TabIndex = 2;
            this.radioButtonLanguage1033.TabStop = true;
            this.radioButtonLanguage1033.Text = "English (US)";
            this.radioButtonLanguage1033.UseVisualStyleBackColor = true;
            this.radioButtonLanguage1033.CheckedChanged += new System.EventHandler(this.radioButtonLanguage1033_CheckedChanged);
            // 
            // buttonDone
            // 
            this.buttonDone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDone.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonDone.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonDone.Image = ((System.Drawing.Image)(resources.GetObject("buttonDone.Image")));
            this.buttonDone.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonDone.Location = new System.Drawing.Point(180, 345);
            this.buttonDone.Name = "buttonDone";
            this.buttonDone.Size = new System.Drawing.Size(125, 29);
            this.buttonDone.TabIndex = 4;
            this.buttonDone.Text = "Close";
            this.buttonDone.UseVisualStyleBackColor = true;
            this.buttonDone.Click += new System.EventHandler(this.buttonDone_Click);
            // 
            // groupBoxOnlineMode
            // 
            this.groupBoxOnlineMode.Controls.Add(this.labelDocumentationHint);
            this.groupBoxOnlineMode.Controls.Add(this.radioButtonConnect);
            this.groupBoxOnlineMode.Controls.Add(this.radioButtonShowLink);
            this.groupBoxOnlineMode.Location = new System.Drawing.Point(23, 134);
            this.groupBoxOnlineMode.Name = "groupBoxOnlineMode";
            this.groupBoxOnlineMode.Size = new System.Drawing.Size(282, 185);
            this.groupBoxOnlineMode.TabIndex = 5;
            this.groupBoxOnlineMode.TabStop = false;
            this.groupBoxOnlineMode.Text = "Tutorial Content";
            // 
            // labelDocumentationHint
            // 
            this.labelDocumentationHint.BackColor = System.Drawing.Color.Orange;
            this.labelDocumentationHint.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelDocumentationHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelDocumentationHint.Location = new System.Drawing.Point(19, 84);
            this.labelDocumentationHint.Name = "labelDocumentationHint";
            this.labelDocumentationHint.Size = new System.Drawing.Size(246, 83);
            this.labelDocumentationHint.TabIndex = 4;
            this.labelDocumentationHint.Text = "The Tutorial application performs a connect to the NetOffice Documentation page o" +
                "ur shows you the link to tutorial documentation.";
            // 
            // radioButtonConnect
            // 
            this.radioButtonConnect.AutoSize = true;
            this.radioButtonConnect.Checked = true;
            this.radioButtonConnect.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonConnect.Location = new System.Drawing.Point(19, 53);
            this.radioButtonConnect.Name = "radioButtonConnect";
            this.radioButtonConnect.Size = new System.Drawing.Size(179, 17);
            this.radioButtonConnect.TabIndex = 3;
            this.radioButtonConnect.TabStop = true;
            this.radioButtonConnect.Text = "Connect to Documentation Page";
            this.radioButtonConnect.UseVisualStyleBackColor = true;
            // 
            // radioButtonShowLink
            // 
            this.radioButtonShowLink.AutoSize = true;
            this.radioButtonShowLink.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonShowLink.Location = new System.Drawing.Point(19, 30);
            this.radioButtonShowLink.Name = "radioButtonShowLink";
            this.radioButtonShowLink.Size = new System.Drawing.Size(182, 17);
            this.radioButtonShowLink.TabIndex = 2;
            this.radioButtonShowLink.Text = "Show Online Documentation Link";
            this.radioButtonShowLink.UseVisualStyleBackColor = true;
            this.radioButtonShowLink.CheckedChanged += new System.EventHandler(this.radioButtonShowLink_CheckedChanged);
            // 
            // checkBoxSaveSettings
            // 
            this.checkBoxSaveSettings.AutoSize = true;
            this.checkBoxSaveSettings.Checked = true;
            this.checkBoxSaveSettings.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSaveSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxSaveSettings.Location = new System.Drawing.Point(42, 345);
            this.checkBoxSaveSettings.Name = "checkBoxSaveSettings";
            this.checkBoxSaveSettings.Size = new System.Drawing.Size(106, 17);
            this.checkBoxSaveSettings.TabIndex = 6;
            this.checkBoxSaveSettings.Text = "Save this settings";
            this.checkBoxSaveSettings.UseVisualStyleBackColor = true;
            this.checkBoxSaveSettings.CheckedChanged += new System.EventHandler(this.checkBoxSaveSettings_CheckedChanged);
            // 
            // FormOptions
            // 
            this.AcceptButton = this.buttonDone;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.CancelButton = this.buttonDone;
            this.ClientSize = new System.Drawing.Size(333, 389);
            this.Controls.Add(this.checkBoxSaveSettings);
            this.Controls.Add(this.groupBoxOnlineMode);
            this.Controls.Add(this.buttonDone);
            this.Controls.Add(this.groupBoxLanguage);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormOptions";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Options";
            this.groupBoxLanguage.ResumeLayout(false);
            this.groupBoxLanguage.PerformLayout();
            this.groupBoxOnlineMode.ResumeLayout(false);
            this.groupBoxOnlineMode.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxLanguage;
        private System.Windows.Forms.RadioButton radioButtonLanguage1031;
        private System.Windows.Forms.RadioButton radioButtonLanguage1033;
        private System.Windows.Forms.Button buttonDone;
        private System.Windows.Forms.GroupBox groupBoxOnlineMode;
        private System.Windows.Forms.RadioButton radioButtonConnect;
        private System.Windows.Forms.RadioButton radioButtonShowLink;
        private System.Windows.Forms.CheckBox checkBoxSaveSettings;
        private System.Windows.Forms.Label labelDocumentationHint;

    }
}
