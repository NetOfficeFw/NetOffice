namespace Example03
{
    partial class FinishDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonClose = new System.Windows.Forms.Button();
            this.buttonOpenDocument = new System.Windows.Forms.Button();
            this.labelMessage = new System.Windows.Forms.Label();
            this.labelDocumentPath = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(243, 79);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(102, 22);
            this.buttonClose.TabIndex = 2;
            this.buttonClose.Text = "Ok";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // buttonOpenDocument
            // 
            this.buttonOpenDocument.Location = new System.Drawing.Point(25, 79);
            this.buttonOpenDocument.Name = "buttonOpenDocument";
            this.buttonOpenDocument.Size = new System.Drawing.Size(102, 22);
            this.buttonOpenDocument.TabIndex = 3;
            this.buttonOpenDocument.Text = "Open Document";
            this.buttonOpenDocument.UseVisualStyleBackColor = true;
            this.buttonOpenDocument.Click += new System.EventHandler(this.buttonOpenDocument_Click);
            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = true;
            this.labelMessage.Location = new System.Drawing.Point(24, 13);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(72, 13);
            this.labelMessage.TabIndex = 4;
            this.labelMessage.Text = "labelMessage";
            // 
            // labelDocumentPath
            // 
            this.labelDocumentPath.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.labelDocumentPath.Location = new System.Drawing.Point(24, 35);
            this.labelDocumentPath.Name = "labelDocumentPath";
            this.labelDocumentPath.Size = new System.Drawing.Size(321, 41);
            this.labelDocumentPath.TabIndex = 5;
            this.labelDocumentPath.Text = "labelDocumentPath";
            // 
            // FinishDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(360, 113);
            this.Controls.Add(this.labelDocumentPath);
            this.Controls.Add(this.labelMessage);
            this.Controls.Add(this.buttonOpenDocument);
            this.Controls.Add(this.buttonClose);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FinishDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Example03";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Button buttonOpenDocument;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.Label labelDocumentPath;

    }
}