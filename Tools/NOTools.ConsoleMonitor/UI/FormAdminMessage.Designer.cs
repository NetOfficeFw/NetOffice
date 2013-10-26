namespace NOTools.ConsoleMonitor
{
    partial class FormAdminMessage
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormAdminMessage));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.pictureBoxAbout = new System.Windows.Forms.PictureBox();
            this.buttonOk = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxAbout)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(85, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(245, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Console Monitor is running with admin permssions. ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(83, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(298, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Applications without admin permission can\'t send notifications.";
            // 
            // pictureBoxAbout
            // 
            this.pictureBoxAbout.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxAbout.Image")));
            this.pictureBoxAbout.Location = new System.Drawing.Point(32, 25);
            this.pictureBoxAbout.Name = "pictureBoxAbout";
            this.pictureBoxAbout.Size = new System.Drawing.Size(33, 33);
            this.pictureBoxAbout.TabIndex = 9;
            this.pictureBoxAbout.TabStop = false;
            // 
            // buttonOk
            // 
            this.buttonOk.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonOk.Location = new System.Drawing.Point(145, 108);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(99, 28);
            this.buttonOk.TabIndex = 10;
            this.buttonOk.Text = "Sad But True";
            this.buttonOk.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.DarkBlue;
            this.label3.Location = new System.Drawing.Point(84, 70);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(124, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "It`s for your own security.";
            // 
            // FormAdminMessage
            // 
            this.AcceptButton = this.buttonOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonOk;
            this.ClientSize = new System.Drawing.Size(426, 162);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.pictureBoxAbout);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormAdminMessage";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Warning";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxAbout)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox pictureBoxAbout;
        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.Label label3;
    }
}