namespace ExcelExamplesCS4
{
    partial class Example09
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Example09));
            this.buttonQuitExample = new System.Windows.Forms.Button();
            this.labelEventLogHeader = new System.Windows.Forms.Label();
            this.textBoxEvents = new System.Windows.Forms.TextBox();
            this.textBoxDescription = new System.Windows.Forms.TextBox();
            this.buttonStartExample = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonQuitExample
            // 
            this.buttonQuitExample.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonQuitExample.Enabled = false;
            this.buttonQuitExample.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonQuitExample.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonQuitExample.Image = ((System.Drawing.Image)(resources.GetObject("buttonQuitExample.Image")));
            this.buttonQuitExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonQuitExample.Location = new System.Drawing.Point(30, 52);
            this.buttonQuitExample.Name = "buttonQuitExample";
            this.buttonQuitExample.Size = new System.Drawing.Size(680, 28);
            this.buttonQuitExample.TabIndex = 21;
            this.buttonQuitExample.Text = "Quit Excel";
            this.buttonQuitExample.UseVisualStyleBackColor = true;
            this.buttonQuitExample.Click += new System.EventHandler(this.buttonQuitExample_Click);
            // 
            // labelEventLogHeader
            // 
            this.labelEventLogHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelEventLogHeader.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelEventLogHeader.Location = new System.Drawing.Point(32, 178);
            this.labelEventLogHeader.Name = "labelEventLogHeader";
            this.labelEventLogHeader.Size = new System.Drawing.Size(679, 22);
            this.labelEventLogHeader.TabIndex = 20;
            this.labelEventLogHeader.Text = "EventLog";
            // 
            // textBoxEvents
            // 
            this.textBoxEvents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxEvents.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.textBoxEvents.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxEvents.Location = new System.Drawing.Point(31, 203);
            this.textBoxEvents.Multiline = true;
            this.textBoxEvents.Name = "textBoxEvents";
            this.textBoxEvents.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxEvents.Size = new System.Drawing.Size(679, 80);
            this.textBoxEvents.TabIndex = 19;
            // 
            // textBoxDescription
            // 
            this.textBoxDescription.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxDescription.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxDescription.Location = new System.Drawing.Point(30, 98);
            this.textBoxDescription.Multiline = true;
            this.textBoxDescription.Name = "textBoxDescription";
            this.textBoxDescription.Size = new System.Drawing.Size(680, 67);
            this.textBoxDescription.TabIndex = 18;
            this.textBoxDescription.Text = resources.GetString("textBoxDescription.Text");
            // 
            // buttonStartExample
            // 
            this.buttonStartExample.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonStartExample.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonStartExample.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStartExample.Image = ((System.Drawing.Image)(resources.GetObject("buttonStartExample.Image")));
            this.buttonStartExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonStartExample.Location = new System.Drawing.Point(31, 18);
            this.buttonStartExample.Name = "buttonStartExample";
            this.buttonStartExample.Size = new System.Drawing.Size(680, 28);
            this.buttonStartExample.TabIndex = 17;
            this.buttonStartExample.Text = "Start Excel";
            this.buttonStartExample.UseVisualStyleBackColor = true;
            this.buttonStartExample.Click += new System.EventHandler(this.buttonStartExample_Click);
            // 
            // Example09
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.buttonQuitExample);
            this.Controls.Add(this.labelEventLogHeader);
            this.Controls.Add(this.textBoxEvents);
            this.Controls.Add(this.textBoxDescription);
            this.Controls.Add(this.buttonStartExample);
            this.Name = "Example09";
            this.Size = new System.Drawing.Size(739, 304);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonQuitExample;
        private System.Windows.Forms.Label labelEventLogHeader;
        private System.Windows.Forms.TextBox textBoxEvents;
        private System.Windows.Forms.TextBox textBoxDescription;
        private System.Windows.Forms.Button buttonStartExample;
    }
}
