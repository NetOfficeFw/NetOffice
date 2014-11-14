namespace NetOffice.DeveloperToolbox
{
    partial class ColorLabel
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
            this.timerEffect = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // timerEffect
            // 
            this.timerEffect.Interval = 50;
            this.timerEffect.Tick += new System.EventHandler(this.timerEffect_Tick);
            // 
            // ColorLabel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "ColorLabel";
            this.Size = new System.Drawing.Size(56, 19);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer timerEffect;
    }
}
