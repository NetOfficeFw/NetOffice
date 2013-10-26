namespace NetOffice.DeveloperToolbox
{
    partial class NameControl
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
            this.labelDescription = new System.Windows.Forms.Label();
            this.textBoxDescription = new System.Windows.Forms.TextBox();
            this.labelClassName = new System.Windows.Forms.Label();
            this.textBoxClassName = new System.Windows.Forms.TextBox();
            this.labelHint = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelDescription
            // 
            this.labelDescription.AutoSize = true;
            this.labelDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelDescription.Location = new System.Drawing.Point(40, 86);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(91, 16);
            this.labelDescription.TabIndex = 17;
            this.labelDescription.Text = "Beschreibung";
            // 
            // textBoxDescription
            // 
            this.textBoxDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxDescription.Location = new System.Drawing.Point(43, 104);
            this.textBoxDescription.Name = "textBoxDescription";
            this.textBoxDescription.Size = new System.Drawing.Size(473, 22);
            this.textBoxDescription.TabIndex = 16;
            this.textBoxDescription.TextChanged += new System.EventHandler(this.textBox_TextChanged);
            // 
            // labelClassName
            // 
            this.labelClassName.AutoSize = true;
            this.labelClassName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelClassName.Location = new System.Drawing.Point(40, 33);
            this.labelClassName.Name = "labelClassName";
            this.labelClassName.Size = new System.Drawing.Size(45, 16);
            this.labelClassName.TabIndex = 15;
            this.labelClassName.Text = "Name";
            // 
            // textBoxClassName
            // 
            this.textBoxClassName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxClassName.Location = new System.Drawing.Point(43, 51);
            this.textBoxClassName.Name = "textBoxClassName";
            this.textBoxClassName.Size = new System.Drawing.Size(189, 22);
            this.textBoxClassName.TabIndex = 14;
            this.textBoxClassName.TextChanged += new System.EventHandler(this.textBox_TextChanged);
            // 
            // labelHint
            // 
            this.labelHint.AutoSize = true;
            this.labelHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelHint.ForeColor = System.Drawing.Color.Red;
            this.labelHint.Location = new System.Drawing.Point(239, 55);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(432, 16);
            this.labelHint.TabIndex = 18;
            this.labelHint.Text = "Es existiert im angegebenen Pfad bereits ein Ordner mit diesem Namen";
            this.labelHint.Visible = false;
            // 
            // NameControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelHint);
            this.Controls.Add(this.labelDescription);
            this.Controls.Add(this.textBoxDescription);
            this.Controls.Add(this.labelClassName);
            this.Controls.Add(this.textBoxClassName);
            this.Name = "NameControl";
            this.Size = new System.Drawing.Size(744, 279);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelDescription;
        private System.Windows.Forms.TextBox textBoxDescription;
        private System.Windows.Forms.Label labelClassName;
        private System.Windows.Forms.TextBox textBoxClassName;
        private System.Windows.Forms.Label labelHint;
    }
}
