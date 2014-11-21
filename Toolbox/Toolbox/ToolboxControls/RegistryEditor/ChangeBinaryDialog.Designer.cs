namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    partial class ChangeBinaryDialog
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
            this.changeBinaryControl1 = new NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor.ChangeBinaryControl();
            this.SuspendLayout();
            // 
            // changeBinaryControl1
            // 
            this.changeBinaryControl1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.changeBinaryControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.changeBinaryControl1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.changeBinaryControl1.Location = new System.Drawing.Point(9, 9);
            this.changeBinaryControl1.Name = "changeBinaryControl1";
            this.changeBinaryControl1.Size = new System.Drawing.Size(534, 432);
            this.changeBinaryControl1.TabIndex = 0;
            this.changeBinaryControl1.Close += new System.EventHandler(this.changeBinaryControl1_Close);
            // 
            // ChangeBinaryDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(552, 450);
            this.Controls.Add(this.changeBinaryControl1);
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ChangeBinaryDialog";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Edit Binary";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ChangeBinaryDialog_KeyDown);
            this.ResumeLayout(false);

        }

        #endregion

        private ChangeBinaryControl changeBinaryControl1;


    }
}
