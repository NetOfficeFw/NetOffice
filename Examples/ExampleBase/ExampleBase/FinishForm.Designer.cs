namespace ExampleBase
{
    partial class FinishForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FinishForm));
            this.labelDocumentPath = new System.Windows.Forms.Label();
            this.labelMessage = new System.Windows.Forms.Label();
            this.buttonOpenDocument = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.checkBoxDeleteDocument = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // labelDocumentPath
            // 
            this.labelDocumentPath.ForeColor = System.Drawing.Color.Black;
            this.labelDocumentPath.Location = new System.Drawing.Point(59, 71);
            this.labelDocumentPath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelDocumentPath.Name = "labelDocumentPath";
            this.labelDocumentPath.Size = new System.Drawing.Size(525, 86);
            this.labelDocumentPath.TabIndex = 9;
            this.labelDocumentPath.Text = "labelDocumentPath";
            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = true;
            this.labelMessage.Font = new System.Drawing.Font("Segoe UI", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.labelMessage.ForeColor = System.Drawing.Color.Black;
            this.labelMessage.Location = new System.Drawing.Point(59, 31);
            this.labelMessage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(91, 17);
            this.labelMessage.TabIndex = 8;
            this.labelMessage.Text = "labelMessage";
            // 
            // buttonOpenDocument
            // 
            this.buttonOpenDocument.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonOpenDocument.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOpenDocument.ForeColor = System.Drawing.Color.Blue;
            this.buttonOpenDocument.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpenDocument.Image")));
            this.buttonOpenDocument.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonOpenDocument.Location = new System.Drawing.Point(57, 162);
            this.buttonOpenDocument.Margin = new System.Windows.Forms.Padding(4);
            this.buttonOpenDocument.Name = "buttonOpenDocument";
            this.buttonOpenDocument.Size = new System.Drawing.Size(177, 31);
            this.buttonOpenDocument.TabIndex = 7;
            this.buttonOpenDocument.Text = "Open Document";
            this.buttonOpenDocument.UseVisualStyleBackColor = true;
            this.buttonOpenDocument.Click += new System.EventHandler(this.buttonOpenDocument_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonClose.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.ForeColor = System.Drawing.Color.Blue;
            this.buttonClose.Image = ((System.Drawing.Image)(resources.GetObject("buttonClose.Image")));
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.Location = new System.Drawing.Point(427, 162);
            this.buttonClose.Margin = new System.Windows.Forms.Padding(4);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(159, 31);
            this.buttonClose.TabIndex = 6;
            this.buttonClose.Text = "Ok";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(27, 33);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(24, 20);
            this.pictureBox1.TabIndex = 10;
            this.pictureBox1.TabStop = false;
            // 
            // checkBoxDeleteDocument
            // 
            this.checkBoxDeleteDocument.AutoSize = true;
            this.checkBoxDeleteDocument.Checked = true;
            this.checkBoxDeleteDocument.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxDeleteDocument.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxDeleteDocument.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxDeleteDocument.Location = new System.Drawing.Point(59, 212);
            this.checkBoxDeleteDocument.Margin = new System.Windows.Forms.Padding(4);
            this.checkBoxDeleteDocument.Name = "checkBoxDeleteDocument";
            this.checkBoxDeleteDocument.Size = new System.Drawing.Size(260, 20);
            this.checkBoxDeleteDocument.TabIndex = 11;
            this.checkBoxDeleteDocument.Text = "Delete document when close this dialog";
            this.checkBoxDeleteDocument.UseVisualStyleBackColor = true;
            // 
            // FinishForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.CancelButton = this.buttonClose;
            this.ClientSize = new System.Drawing.Size(643, 249);
            this.Controls.Add(this.checkBoxDeleteDocument);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.labelDocumentPath);
            this.Controls.Add(this.labelMessage);
            this.Controls.Add(this.buttonOpenDocument);
            this.Controls.Add(this.buttonClose);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FinishForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Finish";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelDocumentPath;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.Button buttonOpenDocument;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox checkBoxDeleteDocument;
    }
}