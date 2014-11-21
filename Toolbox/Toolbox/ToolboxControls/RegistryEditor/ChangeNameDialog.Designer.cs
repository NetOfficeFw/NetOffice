namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
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
            this.changeNameControl1 = new NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor.ChangeNameControl();
            this.SuspendLayout();
            // 
            // changeNameControl1
            // 
            this.changeNameControl1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.changeNameControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.changeNameControl1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.changeNameControl1.Location = new System.Drawing.Point(9, 9);
            this.changeNameControl1.Name = "changeNameControl1";
            this.changeNameControl1.Size = new System.Drawing.Size(349, 148);
            this.changeNameControl1.TabIndex = 0;
            this.changeNameControl1.Close += new System.EventHandler(this.changeNameControl1_Close);
            // 
            // ChangeNameDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(367, 166);
            this.Controls.Add(this.changeNameControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ChangeNameDialog";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Edit Name";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ChangeNameDialog_KeyDown);
            this.ResumeLayout(false);

        }

        #endregion

        private ChangeNameControl changeNameControl1;


    }
}
