namespace Example5
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
            this.labelWorkbookPath = new System.Windows.Forms.Label();
            this.labelMessage = new System.Windows.Forms.Label();
            this.buttonOpenWorkbook = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelWorkbookPath
            // 
            this.labelWorkbookPath.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.labelWorkbookPath.Location = new System.Drawing.Point(11, 34);
            this.labelWorkbookPath.Name = "labelWorkbookPath";
            this.labelWorkbookPath.Size = new System.Drawing.Size(321, 41);
            this.labelWorkbookPath.TabIndex = 13;
            this.labelWorkbookPath.Text = "labelWorkbookPath";
            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = true;
            this.labelMessage.Location = new System.Drawing.Point(11, 12);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(72, 13);
            this.labelMessage.TabIndex = 12;
            this.labelMessage.Text = "labelMessage";
            // 
            // buttonOpenWorkbook
            // 
            this.buttonOpenWorkbook.Location = new System.Drawing.Point(12, 78);
            this.buttonOpenWorkbook.Name = "buttonOpenWorkbook";
            this.buttonOpenWorkbook.Size = new System.Drawing.Size(102, 22);
            this.buttonOpenWorkbook.TabIndex = 11;
            this.buttonOpenWorkbook.Text = "Open Workbook";
            this.buttonOpenWorkbook.UseVisualStyleBackColor = true;
            this.buttonOpenWorkbook.Click += new System.EventHandler(this.buttonOpenWorkbook_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(230, 78);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(102, 22);
            this.buttonClose.TabIndex = 10;
            this.buttonClose.Text = "Ok";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // FinishDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(343, 113);
            this.Controls.Add(this.labelWorkbookPath);
            this.Controls.Add(this.labelMessage);
            this.Controls.Add(this.buttonOpenWorkbook);
            this.Controls.Add(this.buttonClose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FinishDialog";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "FinishDialog";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelWorkbookPath;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.Button buttonOpenWorkbook;
        private System.Windows.Forms.Button buttonClose;

    }
}
