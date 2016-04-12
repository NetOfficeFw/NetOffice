namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    partial class ChangeDWordDialog
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
            this.changeDWORDControl1 = new NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor.ChangeDWORDControl();
            this.SuspendLayout();
            // 
            // changeDWORDControl1
            // 
            this.changeDWORDControl1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.changeDWORDControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.changeDWORDControl1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.changeDWORDControl1.Location = new System.Drawing.Point(9, 9);
            this.changeDWORDControl1.Name = "changeDWORDControl1";
            this.changeDWORDControl1.Size = new System.Drawing.Size(354, 193);
            this.changeDWORDControl1.TabIndex = 0;
            this.changeDWORDControl1.Close += new System.EventHandler(this.changeDWORDControl1_Close);
            // 
            // ChangeDWordDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(372, 211);
            this.Controls.Add(this.changeDWORDControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ChangeDWordDialog";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Edit DWORD Value";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ChangeDWORDDialog_KeyDown);
            this.ResumeLayout(false);

        }

        #endregion

        private ChangeDWORDControl changeDWORDControl1;


    }
}
