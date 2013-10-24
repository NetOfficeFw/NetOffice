namespace NOBuildTools.ReferenceAnalyzer
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
            this.LabelFile = new System.Windows.Forms.Label();
            this.ButttonChooseFile = new System.Windows.Forms.Button();
            this.TextBoxFile = new System.Windows.Forms.TextBox();
            this.ButtonStart = new System.Windows.Forms.Button();
            this.RichTextBoxLog = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // LabelFile
            // 
            this.LabelFile.AutoSize = true;
            this.LabelFile.Location = new System.Drawing.Point(20, 28);
            this.LabelFile.Name = "LabelFile";
            this.LabelFile.Size = new System.Drawing.Size(58, 13);
            this.LabelFile.TabIndex = 6;
            this.LabelFile.Text = "Output File";
            // 
            // ButttonChooseFile
            // 
            this.ButttonChooseFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ButttonChooseFile.Location = new System.Drawing.Point(459, 23);
            this.ButttonChooseFile.Name = "ButttonChooseFile";
            this.ButttonChooseFile.Size = new System.Drawing.Size(40, 22);
            this.ButttonChooseFile.TabIndex = 5;
            this.ButttonChooseFile.Text = "...";
            this.ButttonChooseFile.UseVisualStyleBackColor = true;
            this.ButttonChooseFile.Click += new System.EventHandler(this.ButttonChooseFile_Click);
            // 
            // TextBoxFile
            // 
            this.TextBoxFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxFile.Location = new System.Drawing.Point(83, 25);
            this.TextBoxFile.Name = "TextBoxFile";
            this.TextBoxFile.Size = new System.Drawing.Size(366, 20);
            this.TextBoxFile.TabIndex = 4;
            // 
            // ButtonStart
            // 
            this.ButtonStart.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.ButtonStart.Location = new System.Drawing.Point(23, 59);
            this.ButtonStart.Name = "ButtonStart";
            this.ButtonStart.Size = new System.Drawing.Size(476, 23);
            this.ButtonStart.TabIndex = 8;
            this.ButtonStart.Text = "Start";
            this.ButtonStart.UseVisualStyleBackColor = true;
            this.ButtonStart.Click += new System.EventHandler(this.ButtonStart_Click);
            // 
            // RichTextBoxLog
            // 
            this.RichTextBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.RichTextBoxLog.Location = new System.Drawing.Point(23, 109);
            this.RichTextBoxLog.Name = "RichTextBoxLog";
            this.RichTextBoxLog.ReadOnly = true;
            this.RichTextBoxLog.Size = new System.Drawing.Size(476, 282);
            this.RichTextBoxLog.TabIndex = 7;
            this.RichTextBoxLog.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 416);
            this.Controls.Add(this.ButtonStart);
            this.Controls.Add(this.RichTextBoxLog);
            this.Controls.Add(this.LabelFile);
            this.Controls.Add(this.ButttonChooseFile);
            this.Controls.Add(this.TextBoxFile);
            this.Name = "Form1";
            this.Text = "NOBuildTools.ReferenceAnalyzer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label LabelFile;
        private System.Windows.Forms.Button ButttonChooseFile;
        private System.Windows.Forms.TextBox TextBoxFile;
        private System.Windows.Forms.Button ButtonStart;
        private System.Windows.Forms.RichTextBox RichTextBoxLog;
    }
}

