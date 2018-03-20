namespace NetOffice.DeveloperToolbox.Forms
{
    partial class SelectTicketProviderForm
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
            this.AbortButton = new System.Windows.Forms.Button();
            this.OsdnBox = new System.Windows.Forms.RadioButton();
            this.GithubBox = new System.Windows.Forms.RadioButton();
            this.ProceedButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // AbortButton
            // 
            this.AbortButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.AbortButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.AbortButton.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.AbortButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.AbortButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.AbortButton.ForeColor = System.Drawing.Color.Blue;
            this.AbortButton.Location = new System.Drawing.Point(202, 179);
            this.AbortButton.Margin = new System.Windows.Forms.Padding(4);
            this.AbortButton.Name = "AbortButton";
            this.AbortButton.Size = new System.Drawing.Size(124, 31);
            this.AbortButton.TabIndex = 30;
            this.AbortButton.Text = "Cancel";
            this.AbortButton.UseVisualStyleBackColor = true;
            this.AbortButton.Click += new System.EventHandler(this.AbortButton_Click);
            // 
            // OsdnBox
            // 
            this.OsdnBox.AutoSize = true;
            this.OsdnBox.Checked = true;
            this.OsdnBox.Location = new System.Drawing.Point(56, 57);
            this.OsdnBox.Margin = new System.Windows.Forms.Padding(4);
            this.OsdnBox.Name = "OsdnBox";
            this.OsdnBox.Size = new System.Drawing.Size(65, 20);
            this.OsdnBox.TabIndex = 31;
            this.OsdnBox.TabStop = true;
            this.OsdnBox.Tag = "https://osdn.net/projects/netoffice/ticket";
            this.OsdnBox.Text = "OSDN";
            this.OsdnBox.UseVisualStyleBackColor = true;
            // 
            // GithubBox
            // 
            this.GithubBox.AutoSize = true;
            this.GithubBox.Location = new System.Drawing.Point(56, 100);
            this.GithubBox.Margin = new System.Windows.Forms.Padding(4);
            this.GithubBox.Name = "GithubBox";
            this.GithubBox.Size = new System.Drawing.Size(64, 20);
            this.GithubBox.TabIndex = 32;
            this.GithubBox.Tag = "https://github.com/NetOfficeFw/NetOffice/issues";
            this.GithubBox.Text = "Github";
            this.GithubBox.UseVisualStyleBackColor = true;
            // 
            // ProceedButton
            // 
            this.ProceedButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ProceedButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.ProceedButton.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.ProceedButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ProceedButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.ProceedButton.ForeColor = System.Drawing.Color.Blue;
            this.ProceedButton.Location = new System.Drawing.Point(38, 179);
            this.ProceedButton.Margin = new System.Windows.Forms.Padding(4);
            this.ProceedButton.Name = "ProceedButton";
            this.ProceedButton.Size = new System.Drawing.Size(124, 31);
            this.ProceedButton.TabIndex = 33;
            this.ProceedButton.Text = "OK";
            this.ProceedButton.UseVisualStyleBackColor = true;
            this.ProceedButton.Click += new System.EventHandler(this.ProceedButton_Click);
            // 
            // SelectTicketProviderForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.CancelButton = this.AbortButton;
            this.ClientSize = new System.Drawing.Size(380, 248);
            this.Controls.Add(this.ProceedButton);
            this.Controls.Add(this.GithubBox);
            this.Controls.Add(this.OsdnBox);
            this.Controls.Add(this.AbortButton);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SelectTicketProviderForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Ticket Provider";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button AbortButton;
        private System.Windows.Forms.RadioButton OsdnBox;
        private System.Windows.Forms.RadioButton GithubBox;
        private System.Windows.Forms.Button ProceedButton;
    }
}