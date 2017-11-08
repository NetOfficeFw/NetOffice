namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    partial class LoadControl
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
            this.labelUserCaption = new System.Windows.Forms.Label();
            this.comboBoxLoadBehavior = new System.Windows.Forms.ComboBox();
            this.labelLoadCaption = new System.Windows.Forms.Label();
            this.radioButtonLocalMachine = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.radioButtonCurrentUser = new NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton();
            this.SuspendLayout();
            // 
            // labelUserCaption
            // 
            this.labelUserCaption.AutoSize = true;
            this.labelUserCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelUserCaption.ForeColor = System.Drawing.Color.Black;
            this.labelUserCaption.Location = new System.Drawing.Point(40, 33);
            this.labelUserCaption.Name = "labelUserCaption";
            this.labelUserCaption.Size = new System.Drawing.Size(317, 16);
            this.labelUserCaption.TabIndex = 21;
            this.labelUserCaption.Text = "Choose addin is available for all users or not";
            // 
            // comboBoxLoadBehavior
            // 
            this.comboBoxLoadBehavior.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxLoadBehavior.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBoxLoadBehavior.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxLoadBehavior.FormattingEnabled = true;
            this.comboBoxLoadBehavior.Items.AddRange(new object[] {
            "3   = Load at startup",
            "2   = Load on demand",
            "1   = Do not load automatically",
            "16 = Load first time at startup, then load on demand"});
            this.comboBoxLoadBehavior.Location = new System.Drawing.Point(46, 148);
            this.comboBoxLoadBehavior.Name = "comboBoxLoadBehavior";
            this.comboBoxLoadBehavior.Size = new System.Drawing.Size(477, 25);
            this.comboBoxLoadBehavior.TabIndex = 20;
            this.comboBoxLoadBehavior.SelectedIndexChanged += new System.EventHandler(this.comboBoxLoadBehavior_SelectedIndexChanged);
            // 
            // labelLoadCaption
            // 
            this.labelLoadCaption.AutoSize = true;
            this.labelLoadCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelLoadCaption.ForeColor = System.Drawing.Color.Black;
            this.labelLoadCaption.Location = new System.Drawing.Point(43, 119);
            this.labelLoadCaption.Name = "labelLoadCaption";
            this.labelLoadCaption.Size = new System.Drawing.Size(230, 16);
            this.labelLoadCaption.TabIndex = 19;
            this.labelLoadCaption.Text = "Decide when it has to be loaded";
            // 
            // radioButtonLocalMachine
            // 
            this.radioButtonLocalMachine.AutoSize = true;
            this.radioButtonLocalMachine.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonLocalMachine.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonLocalMachine.ForeColor = System.Drawing.SystemColors.ControlText;
            this.radioButtonLocalMachine.Location = new System.Drawing.Point(46, 81);
            this.radioButtonLocalMachine.Name = "radioButtonLocalMachine";
            this.radioButtonLocalMachine.Size = new System.Drawing.Size(79, 20);
            this.radioButtonLocalMachine.TabIndex = 18;
            this.radioButtonLocalMachine.Text = "All Users";
            this.radioButtonLocalMachine.UseVisualStyleBackColor = true;
            this.radioButtonLocalMachine.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // radioButtonCurrentUser
            // 
            this.radioButtonCurrentUser.AutoSize = true;
            this.radioButtonCurrentUser.Checked = true;
            this.radioButtonCurrentUser.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonCurrentUser.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonCurrentUser.ForeColor = System.Drawing.Color.Blue;
            this.radioButtonCurrentUser.Location = new System.Drawing.Point(46, 58);
            this.radioButtonCurrentUser.Name = "radioButtonCurrentUser";
            this.radioButtonCurrentUser.Size = new System.Drawing.Size(99, 20);
            this.radioButtonCurrentUser.TabIndex = 17;
            this.radioButtonCurrentUser.TabStop = true;
            this.radioButtonCurrentUser.Text = "Current User";
            this.radioButtonCurrentUser.UseVisualStyleBackColor = true;
            this.radioButtonCurrentUser.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // LoadControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.labelUserCaption);
            this.Controls.Add(this.comboBoxLoadBehavior);
            this.Controls.Add(this.labelLoadCaption);
            this.Controls.Add(this.radioButtonLocalMachine);
            this.Controls.Add(this.radioButtonCurrentUser);
            this.Name = "LoadControl";
            this.Size = new System.Drawing.Size(611, 285);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelUserCaption;
        private System.Windows.Forms.ComboBox comboBoxLoadBehavior;
        private System.Windows.Forms.Label labelLoadCaption;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonLocalMachine;
        private NetOffice.DeveloperToolbox.Controls.Radio.GlowRadioButton radioButtonCurrentUser;
    }
}
