namespace NetOffice.ProjectWizard
{
    partial class AddinLoadControl
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
            this.radioButtonCurrentUser = new System.Windows.Forms.RadioButton();
            this.radioButtonLocalMachine = new System.Windows.Forms.RadioButton();
            this.labelLoadCaption = new System.Windows.Forms.Label();
            this.comboBoxLoadBehavior = new System.Windows.Forms.ComboBox();
            this.labelUserCaption = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // radioButtonCurrentUser
            // 
            this.radioButtonCurrentUser.AutoSize = true;
            this.radioButtonCurrentUser.Checked = true;
            this.radioButtonCurrentUser.Location = new System.Drawing.Point(33, 57);
            this.radioButtonCurrentUser.Name = "radioButtonCurrentUser";
            this.radioButtonCurrentUser.Size = new System.Drawing.Size(193, 17);
            this.radioButtonCurrentUser.TabIndex = 12;
            this.radioButtonCurrentUser.TabStop = true;
            this.radioButtonCurrentUser.Text = "Nur für den angemeldeten Benutzer";
            this.radioButtonCurrentUser.UseVisualStyleBackColor = true;
            this.radioButtonCurrentUser.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // radioButtonLocalMachine
            // 
            this.radioButtonLocalMachine.AutoSize = true;
            this.radioButtonLocalMachine.Location = new System.Drawing.Point(33, 80);
            this.radioButtonLocalMachine.Name = "radioButtonLocalMachine";
            this.radioButtonLocalMachine.Size = new System.Drawing.Size(104, 17);
            this.radioButtonLocalMachine.TabIndex = 13;
            this.radioButtonLocalMachine.Text = "Für alle Benutzer";
            this.radioButtonLocalMachine.UseVisualStyleBackColor = true;
            this.radioButtonLocalMachine.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // labelLoadCaption
            // 
            this.labelLoadCaption.AutoSize = true;
            this.labelLoadCaption.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelLoadCaption.Location = new System.Drawing.Point(30, 118);
            this.labelLoadCaption.Name = "labelLoadCaption";
            this.labelLoadCaption.Size = new System.Drawing.Size(292, 15);
            this.labelLoadCaption.TabIndex = 14;
            this.labelLoadCaption.Text = "Wann soll Ihr Automations Addin geladen werden?";
            // 
            // comboBoxLoadBehavior
            // 
            this.comboBoxLoadBehavior.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxLoadBehavior.FormattingEnabled = true;
            this.comboBoxLoadBehavior.Items.AddRange(new object[] {
            "3   = Beim Start der Office Anwendung automatisch laden",
            "2   = Bei Bedarf laden",
            "1   = Nicht automatisch laden",
            "16 = Beim ersten Start automatisch laden, danach bei Bedarf laden"});
            this.comboBoxLoadBehavior.Location = new System.Drawing.Point(33, 147);
            this.comboBoxLoadBehavior.Name = "comboBoxLoadBehavior";
            this.comboBoxLoadBehavior.Size = new System.Drawing.Size(353, 21);
            this.comboBoxLoadBehavior.TabIndex = 15;
            this.comboBoxLoadBehavior.SelectedIndexChanged += new System.EventHandler(this.comboBoxLoadBehavior_SelectedIndexChanged);
            // 
            // labelUserCaption
            // 
            this.labelUserCaption.AutoSize = true;
            this.labelUserCaption.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelUserCaption.Location = new System.Drawing.Point(30, 30);
            this.labelUserCaption.Name = "labelUserCaption";
            this.labelUserCaption.Size = new System.Drawing.Size(295, 15);
            this.labelUserCaption.TabIndex = 16;
            this.labelUserCaption.Text = "Für wen soll Ihr Automations Addin verfügbar sein?";
            // 
            // AddinLoadControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelUserCaption);
            this.Controls.Add(this.comboBoxLoadBehavior);
            this.Controls.Add(this.labelLoadCaption);
            this.Controls.Add(this.radioButtonLocalMachine);
            this.Controls.Add(this.radioButtonCurrentUser);
            this.Name = "AddinLoadControl";
            this.Size = new System.Drawing.Size(520, 200);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton radioButtonCurrentUser;
        private System.Windows.Forms.RadioButton radioButtonLocalMachine;
        private System.Windows.Forms.Label labelLoadCaption;
        private System.Windows.Forms.ComboBox comboBoxLoadBehavior;
        private System.Windows.Forms.Label labelUserCaption;
    }
}
