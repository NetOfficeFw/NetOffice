namespace Example1
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
            this.buttonOpenWorkbook = new System.Windows.Forms.Button();
            this.labelMessage = new System.Windows.Forms.Label();
            this.labelWorkbookPath = new System.Windows.Forms.Label();
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
            // buttonOpenWorkbook
            // 
            this.buttonOpenWorkbook.Location = new System.Drawing.Point(25, 79);
            this.buttonOpenWorkbook.Name = "buttonOpenWorkbook";
            this.buttonOpenWorkbook.Size = new System.Drawing.Size(102, 22);
            this.buttonOpenWorkbook.TabIndex = 3;
            this.buttonOpenWorkbook.Text = "Open Workbook";
            this.buttonOpenWorkbook.UseVisualStyleBackColor = true;
            this.buttonOpenWorkbook.Click += new System.EventHandler(this.buttonOpenWorkbook_Click);
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
            // labelWorkbookPath
            // 
            this.labelWorkbookPath.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.labelWorkbookPath.Location = new System.Drawing.Point(24, 35);
            this.labelWorkbookPath.Name = "labelWorkbookPath";
            this.labelWorkbookPath.Size = new System.Drawing.Size(321, 41);
            this.labelWorkbookPath.TabIndex = 5;
            this.labelWorkbookPath.Text = "labelWorkbookPath";
            // 
            // FinishDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(360, 113);
            this.Controls.Add(this.labelWorkbookPath);
            this.Controls.Add(this.labelMessage);
            this.Controls.Add(this.buttonOpenWorkbook);
            this.Controls.Add(this.buttonClose);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FinishDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Example1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Button buttonOpenWorkbook;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.Label labelWorkbookPath;

    }
}