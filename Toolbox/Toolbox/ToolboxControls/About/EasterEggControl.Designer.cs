namespace NetOffice.DeveloperToolbox.ToolboxControls.About
{
    partial class EasterEggControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EasterEggControl));
            this.pictureBoxWait = new System.Windows.Forms.PictureBox();
            this.pictureBoxGernot = new System.Windows.Forms.PictureBox();
            this.panelMessage = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxWait)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxGernot)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBoxWait
            // 
            this.pictureBoxWait.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxWait.BackgroundImage")));
            this.pictureBoxWait.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxWait.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.pictureBoxWait.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBoxWait.Location = new System.Drawing.Point(0, 0);
            this.pictureBoxWait.Name = "pictureBoxWait";
            this.pictureBoxWait.Size = new System.Drawing.Size(936, 601);
            this.pictureBoxWait.TabIndex = 16;
            this.pictureBoxWait.TabStop = false;
            this.pictureBoxWait.Visible = false;
            // 
            // pictureBoxGernot
            // 
            this.pictureBoxGernot.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxGernot.BackgroundImage")));
            this.pictureBoxGernot.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxGernot.Location = new System.Drawing.Point(180, 141);
            this.pictureBoxGernot.Name = "pictureBoxGernot";
            this.pictureBoxGernot.Size = new System.Drawing.Size(553, 396);
            this.pictureBoxGernot.TabIndex = 17;
            this.pictureBoxGernot.TabStop = false;
            this.pictureBoxGernot.Visible = false;
            // 
            // panelMessage
            // 
            this.panelMessage.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panelMessage.Location = new System.Drawing.Point(180, 41);
            this.panelMessage.Name = "panelMessage";
            this.panelMessage.Size = new System.Drawing.Size(553, 100);
            this.panelMessage.TabIndex = 18;
            // 
            // EasterEggControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.panelMessage);
            this.Controls.Add(this.pictureBoxGernot);
            this.Controls.Add(this.pictureBoxWait);
            this.Name = "EasterEggControl";
            this.Size = new System.Drawing.Size(936, 601);
            this.Resize += new System.EventHandler(this.EasterEggControl_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxWait)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxGernot)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBoxWait;
        private System.Windows.Forms.PictureBox pictureBoxGernot;
        private System.Windows.Forms.Panel panelMessage;
    }
}
