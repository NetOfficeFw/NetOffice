namespace NetOffice.DeveloperToolbox.Forms
{
    partial class UriTargetForm
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
            this.components = new System.ComponentModel.Container();
            this.HeaderPanel = new System.Windows.Forms.Panel();
            this.OsdnRadioButton = new System.Windows.Forms.RadioButton();
            this.GithubRadioButton = new System.Windows.Forms.RadioButton();
            this.ProceedButton = new System.Windows.Forms.Button();
            this.AbortButton = new System.Windows.Forms.Button();
            this.GithubLinkLabel = new System.Windows.Forms.Label();
            this.OsdnLinkLabel = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.LogonRequiredLabel = new System.Windows.Forms.Label();
            this.LinkContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.HeaderPanel.SuspendLayout();
            this.LinkContextMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // HeaderPanel
            // 
            this.HeaderPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.HeaderPanel.BackColor = System.Drawing.Color.White;
            this.HeaderPanel.Controls.Add(this.label4);
            this.HeaderPanel.Controls.Add(this.label3);
            this.HeaderPanel.Location = new System.Drawing.Point(1, 2);
            this.HeaderPanel.Name = "HeaderPanel";
            this.HeaderPanel.Size = new System.Drawing.Size(302, 65);
            this.HeaderPanel.TabIndex = 0;
            // 
            // OsdnRadioButton
            // 
            this.OsdnRadioButton.AutoSize = true;
            this.OsdnRadioButton.Checked = true;
            this.OsdnRadioButton.Location = new System.Drawing.Point(35, 96);
            this.OsdnRadioButton.Name = "OsdnRadioButton";
            this.OsdnRadioButton.Size = new System.Drawing.Size(66, 17);
            this.OsdnRadioButton.TabIndex = 1;
            this.OsdnRadioButton.TabStop = true;
            this.OsdnRadioButton.Tag = "https://osdn.net/ticket/newticket.php?group_id=10754";
            this.OsdnRadioButton.Text = "osdn.net";
            this.OsdnRadioButton.UseVisualStyleBackColor = true;
            // 
            // GithubRadioButton
            // 
            this.GithubRadioButton.AutoSize = true;
            this.GithubRadioButton.Location = new System.Drawing.Point(35, 137);
            this.GithubRadioButton.Name = "GithubRadioButton";
            this.GithubRadioButton.Size = new System.Drawing.Size(77, 17);
            this.GithubRadioButton.TabIndex = 2;
            this.GithubRadioButton.Tag = "https://github.com/NetOfficeFw/NetOffice/issues";
            this.GithubRadioButton.Text = "github.com";
            this.GithubRadioButton.UseVisualStyleBackColor = true;
            // 
            // ProceedButton
            // 
            this.ProceedButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ProceedButton.Location = new System.Drawing.Point(105, 206);
            this.ProceedButton.Name = "ProceedButton";
            this.ProceedButton.Size = new System.Drawing.Size(75, 23);
            this.ProceedButton.TabIndex = 3;
            this.ProceedButton.Text = "Proceed";
            this.ProceedButton.UseVisualStyleBackColor = true;
            // 
            // AbortButton
            // 
            this.AbortButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.AbortButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.AbortButton.Location = new System.Drawing.Point(207, 206);
            this.AbortButton.Name = "AbortButton";
            this.AbortButton.Size = new System.Drawing.Size(75, 23);
            this.AbortButton.TabIndex = 4;
            this.AbortButton.Text = "Abort";
            this.AbortButton.UseVisualStyleBackColor = true;
            // 
            // GithubLinkLabel
            // 
            this.GithubLinkLabel.AutoSize = true;
            this.GithubLinkLabel.ForeColor = System.Drawing.Color.Gray;
            this.GithubLinkLabel.Location = new System.Drawing.Point(123, 140);
            this.GithubLinkLabel.Name = "GithubLinkLabel";
            this.GithubLinkLabel.Size = new System.Drawing.Size(159, 13);
            this.GithubLinkLabel.TabIndex = 5;
            this.GithubLinkLabel.Tag = "https://github.com/NetOfficeFw/NetOffice/issues";
            this.GithubLinkLabel.Text = "https://github.com/.../NetOffice";
            // 
            // OsdnLinkLabel
            // 
            this.OsdnLinkLabel.AutoSize = true;
            this.OsdnLinkLabel.ForeColor = System.Drawing.Color.Gray;
            this.OsdnLinkLabel.Location = new System.Drawing.Point(123, 98);
            this.OsdnLinkLabel.Name = "OsdnLinkLabel";
            this.OsdnLinkLabel.Size = new System.Drawing.Size(170, 13);
            this.OsdnLinkLabel.TabIndex = 6;
            this.OsdnLinkLabel.Tag = "https://osdn.net/ticket/newticket.php?group_id=10754";
            this.OsdnLinkLabel.Text = "https://osdn.net/.../newticket.php";
            this.OsdnLinkLabel.Click += new System.EventHandler(this.OsdnLinkLabel_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label3.Location = new System.Drawing.Point(17, 14);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(134, 16);
            this.label3.TabIndex = 0;
            this.label3.Text = "Make your choice.";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.Gray;
            this.label4.Location = new System.Drawing.Point(35, 36);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(182, 13);
            this.label4.TabIndex = 7;
            this.label4.Tag = "";
            this.label4.Text = "Answers on github.com may delayed.";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(95, 97);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(11, 13);
            this.label5.TabIndex = 7;
            this.label5.Text = "*";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.Red;
            this.label6.Location = new System.Drawing.Point(106, 139);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(11, 13);
            this.label6.TabIndex = 8;
            this.label6.Text = "*";
            // 
            // LogonRequiredLabel
            // 
            this.LogonRequiredLabel.AutoSize = true;
            this.LogonRequiredLabel.ForeColor = System.Drawing.Color.Red;
            this.LogonRequiredLabel.Location = new System.Drawing.Point(34, 172);
            this.LogonRequiredLabel.Name = "LogonRequiredLabel";
            this.LogonRequiredLabel.Size = new System.Drawing.Size(85, 13);
            this.LogonRequiredLabel.TabIndex = 9;
            this.LogonRequiredLabel.Text = "* Logon required";
            // 
            // LinkContextMenu
            // 
            this.LinkContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.LinkContextMenu.Name = "LinkContextMenu";
            this.LinkContextMenu.Size = new System.Drawing.Size(177, 48);
            this.LinkContextMenu.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.LinkContextMenu_ItemClicked);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(176, 22);
            this.toolStripMenuItem1.Text = "Copy Link Location";
            // 
            // UriTargetForm
            // 
            this.AcceptButton = this.ProceedButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.CancelButton = this.AbortButton;
            this.ClientSize = new System.Drawing.Size(305, 252);
            this.Controls.Add(this.LogonRequiredLabel);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.OsdnLinkLabel);
            this.Controls.Add(this.GithubLinkLabel);
            this.Controls.Add(this.AbortButton);
            this.Controls.Add(this.ProceedButton);
            this.Controls.Add(this.GithubRadioButton);
            this.Controls.Add(this.HeaderPanel);
            this.Controls.Add(this.OsdnRadioButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UriTargetForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Platform Selection";
            this.Click += new System.EventHandler(this.UriTargetForm_Click);
            this.HeaderPanel.ResumeLayout(false);
            this.HeaderPanel.PerformLayout();
            this.LinkContextMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel HeaderPanel;
        private System.Windows.Forms.RadioButton OsdnRadioButton;
        private System.Windows.Forms.RadioButton GithubRadioButton;
        private System.Windows.Forms.Button ProceedButton;
        private System.Windows.Forms.Button AbortButton;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label GithubLinkLabel;
        private System.Windows.Forms.Label OsdnLinkLabel;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label LogonRequiredLabel;
        private System.Windows.Forms.ContextMenuStrip LinkContextMenu;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
    }
}