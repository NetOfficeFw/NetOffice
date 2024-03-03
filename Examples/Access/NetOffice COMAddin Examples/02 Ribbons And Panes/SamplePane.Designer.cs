namespace Access02AddinCS4
{
    partial class SamplePane
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
            this.components = new System.ComponentModel.Container();
            this.UsageTimer = new System.Windows.Forms.Timer(this.components);
            this.UsageLabel = new System.Windows.Forms.Label();
            this.UsageBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // UsageTimer
            // 
            this.UsageTimer.Interval = 400;
            this.UsageTimer.Tick += new System.EventHandler(this.UsageTimer_Tick);
            // 
            // UsageLabel
            // 
            this.UsageLabel.AutoSize = true;
            this.UsageLabel.BackColor = System.Drawing.Color.Transparent;
            this.UsageLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.UsageLabel.ForeColor = System.Drawing.Color.Blue;
            this.UsageLabel.Location = new System.Drawing.Point(126, 8);
            this.UsageLabel.Name = "UsageLabel";
            this.UsageLabel.Size = new System.Drawing.Size(48, 13);
            this.UsageLabel.TabIndex = 7;
            this.UsageLabel.Text = "<Empty>";
            this.UsageLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // UsageBar
            // 
            this.UsageBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UsageBar.Location = new System.Drawing.Point(0, 0);
            this.UsageBar.Name = "UsageBar";
            this.UsageBar.Size = new System.Drawing.Size(300, 30);
            this.UsageBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.UsageBar.TabIndex = 8;
            // 
            // SamplePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.UsageLabel);
            this.Controls.Add(this.UsageBar);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Blue;
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "SamplePane";
            this.Size = new System.Drawing.Size(300, 30);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer UsageTimer;
        private System.Windows.Forms.Label UsageLabel;
        private System.Windows.Forms.ProgressBar UsageBar;
    }
}
