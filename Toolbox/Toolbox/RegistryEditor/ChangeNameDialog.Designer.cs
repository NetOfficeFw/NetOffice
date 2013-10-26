namespace NetOffice.DeveloperToolbox.RegistryEditor
{
    partial class ChangeNameDialog
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
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonAbort = new System.Windows.Forms.Button();
            this.textBoxValue = new System.Windows.Forms.TextBox();
            this.labelNewName = new System.Windows.Forms.Label();
            this.textBoxName = new System.Windows.Forms.TextBox();
            this.labelName = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.Location = new System.Drawing.Point(176, 105);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(74, 22);
            this.buttonOK.TabIndex = 16;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonAbort
            // 
            this.buttonAbort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonAbort.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonAbort.Location = new System.Drawing.Point(256, 105);
            this.buttonAbort.Name = "buttonAbort";
            this.buttonAbort.Size = new System.Drawing.Size(74, 22);
            this.buttonAbort.TabIndex = 15;
            this.buttonAbort.Text = "Abbrechen";
            this.buttonAbort.UseVisualStyleBackColor = true;
            this.buttonAbort.Click += new System.EventHandler(this.buttonAbort_Click);
            // 
            // textBoxValue
            // 
            this.textBoxValue.Location = new System.Drawing.Point(12, 69);
            this.textBoxValue.Name = "textBoxValue";
            this.textBoxValue.Size = new System.Drawing.Size(320, 20);
            this.textBoxValue.TabIndex = 1;
            // 
            // labelNewName
            // 
            this.labelNewName.AutoSize = true;
            this.labelNewName.Location = new System.Drawing.Point(9, 53);
            this.labelNewName.Name = "labelNewName";
            this.labelNewName.Size = new System.Drawing.Size(67, 13);
            this.labelNewName.TabIndex = 19;
            this.labelNewName.Text = "Neuer Name";
            // 
            // textBoxName
            // 
            this.textBoxName.Location = new System.Drawing.Point(12, 24);
            this.textBoxName.Name = "textBoxName";
            this.textBoxName.ReadOnly = true;
            this.textBoxName.Size = new System.Drawing.Size(319, 20);
            this.textBoxName.TabIndex = 20;
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Location = new System.Drawing.Point(9, 8);
            this.labelName.Name = "labelName";
            this.labelName.Size = new System.Drawing.Size(35, 13);
            this.labelName.TabIndex = 17;
            this.labelName.Text = "Name";
            // 
            // ChangeNameDialog
            // 
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonAbort;
            this.ClientSize = new System.Drawing.Size(342, 139);
            this.Controls.Add(this.textBoxValue);
            this.Controls.Add(this.labelNewName);
            this.Controls.Add(this.textBoxName);
            this.Controls.Add(this.labelName);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonAbort);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ChangeNameDialog";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Name bearbeiten";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonAbort;
        private System.Windows.Forms.TextBox textBoxValue;
        private System.Windows.Forms.Label labelNewName;
        private System.Windows.Forms.TextBox textBoxName;
        private System.Windows.Forms.Label labelName;

    }
}
