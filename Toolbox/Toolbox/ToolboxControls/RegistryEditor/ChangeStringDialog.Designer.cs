namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    partial class ChangeStringDialog
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
            this.changeStringControl1 = new NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor.ChangeStringControl();
            this.SuspendLayout();
            // 
            // changeStringControl1
            // 
            this.changeStringControl1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.changeStringControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.changeStringControl1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.changeStringControl1.Location = new System.Drawing.Point(9, 9);
            this.changeStringControl1.Name = "changeStringControl1";
            this.changeStringControl1.Size = new System.Drawing.Size(345, 137);
            this.changeStringControl1.TabIndex = 0;
            this.changeStringControl1.Close += new System.EventHandler(this.changeStringControl1_Close);
            // 
            // ChangeStringDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(363, 155);
            this.Controls.Add(this.changeStringControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ChangeStringDialog";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Edit String";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ChangeStringDialog_KeyDown);
            this.ResumeLayout(false);

        }

        #endregion

        private ChangeStringControl changeStringControl1;


    }
}
