namespace NOToolsTests.CSharpTextEditor1
{
    partial class Form1
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

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.codeEditorControl1 = new NOTools.CSharpTextEditor.CodeEditorControl();
            this.SuspendLayout();
            // 
            // codeEditorControl1
            // 
            this.codeEditorControl1.BackColor = System.Drawing.Color.Black;
            this.codeEditorControl1.Location = new System.Drawing.Point(0, 2);
            this.codeEditorControl1.MinimumSize = new System.Drawing.Size(400, 300);
            this.codeEditorControl1.Name = "codeEditorControl1";
            this.codeEditorControl1.PersistencePath = null;
            this.codeEditorControl1.ShowLineNumbers = true;
            this.codeEditorControl1.Size = new System.Drawing.Size(562, 317);
            this.codeEditorControl1.TabIndex = 0;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 393);
            this.Controls.Add(this.codeEditorControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private NOTools.CSharpTextEditor.CodeEditorControl codeEditorControl1;
    }
}

