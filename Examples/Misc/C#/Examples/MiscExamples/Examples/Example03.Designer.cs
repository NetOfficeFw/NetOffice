namespace MiscExamplesCS4
{
    partial class Example03
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Example03));
            this.textBoxDescription = new System.Windows.Forms.TextBox();
            this.buttonStartExample = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.labelEventLogHeader = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBoxDescription
            // 
            this.textBoxDescription.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxDescription.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxDescription.Location = new System.Drawing.Point(28, 72);
            this.textBoxDescription.Multiline = true;
            this.textBoxDescription.Name = "textBoxDescription";
            this.textBoxDescription.Size = new System.Drawing.Size(679, 27);
            this.textBoxDescription.TabIndex = 19;
            this.textBoxDescription.Text = "This example shows you how to access a running Outlook application.";
            // 
            // buttonStartExample
            // 
            this.buttonStartExample.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonStartExample.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStartExample.Image = ((System.Drawing.Image)(resources.GetObject("buttonStartExample.Image")));
            this.buttonStartExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonStartExample.Location = new System.Drawing.Point(28, 23);
            this.buttonStartExample.Name = "buttonStartExample";
            this.buttonStartExample.Size = new System.Drawing.Size(679, 28);
            this.buttonStartExample.TabIndex = 18;
            this.buttonStartExample.Text = "Start example";
            this.buttonStartExample.UseVisualStyleBackColor = true;
            this.buttonStartExample.Click += new System.EventHandler(this.buttonStartExample_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.textBoxLog.Location = new System.Drawing.Point(28, 141);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.Size = new System.Drawing.Size(679, 139);
            this.textBoxLog.TabIndex = 20;
            // 
            // labelEventLogHeader
            // 
            this.labelEventLogHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelEventLogHeader.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelEventLogHeader.Location = new System.Drawing.Point(28, 116);
            this.labelEventLogHeader.Name = "labelEventLogHeader";
            this.labelEventLogHeader.Size = new System.Drawing.Size(679, 22);
            this.labelEventLogHeader.TabIndex = 21;
            this.labelEventLogHeader.Text = "EventLog";
            // 
            // Example03
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelEventLogHeader);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.textBoxDescription);
            this.Controls.Add(this.buttonStartExample);
            this.Name = "Example03";
            this.Size = new System.Drawing.Size(739, 304);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxDescription;
        private System.Windows.Forms.Button buttonStartExample;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.Label labelEventLogHeader;
    }
}
