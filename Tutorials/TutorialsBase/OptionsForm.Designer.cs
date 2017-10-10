namespace TutorialsBase
{
    partial class OptionsForm
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OptionsForm));
            this.buttonDone = new System.Windows.Forms.Button();
            this.groupBoxOnlineMode = new System.Windows.Forms.GroupBox();
            this.radioButtonConnect = new System.Windows.Forms.RadioButton();
            this.radioButtonShowLink = new System.Windows.Forms.RadioButton();
            this.groupBoxOnlineMode.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonDone
            // 
            this.buttonDone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDone.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonDone.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonDone.ForeColor = System.Drawing.Color.Blue;
            this.buttonDone.Image = ((System.Drawing.Image)(resources.GetObject("buttonDone.Image")));
            this.buttonDone.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonDone.Location = new System.Drawing.Point(276, 150);
            this.buttonDone.Name = "buttonDone";
            this.buttonDone.Size = new System.Drawing.Size(125, 29);
            this.buttonDone.TabIndex = 4;
            this.buttonDone.Text = "Close";
            this.buttonDone.UseVisualStyleBackColor = true;
            this.buttonDone.Click += new System.EventHandler(this.buttonDone_Click);
            // 
            // groupBoxOnlineMode
            // 
            this.groupBoxOnlineMode.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxOnlineMode.Controls.Add(this.radioButtonConnect);
            this.groupBoxOnlineMode.Controls.Add(this.radioButtonShowLink);
            this.groupBoxOnlineMode.ForeColor = System.Drawing.Color.Black;
            this.groupBoxOnlineMode.Location = new System.Drawing.Point(22, 25);
            this.groupBoxOnlineMode.Name = "groupBoxOnlineMode";
            this.groupBoxOnlineMode.Size = new System.Drawing.Size(379, 105);
            this.groupBoxOnlineMode.TabIndex = 5;
            this.groupBoxOnlineMode.TabStop = false;
            this.groupBoxOnlineMode.Text = "Tutorial Content";
            // 
            // radioButtonConnect
            // 
            this.radioButtonConnect.AutoSize = true;
            this.radioButtonConnect.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonConnect.ForeColor = System.Drawing.Color.Black;
            this.radioButtonConnect.Location = new System.Drawing.Point(19, 60);
            this.radioButtonConnect.Name = "radioButtonConnect";
            this.radioButtonConnect.Size = new System.Drawing.Size(179, 17);
            this.radioButtonConnect.TabIndex = 3;
            this.radioButtonConnect.Text = "Connect to Documentation Page";
            this.radioButtonConnect.UseVisualStyleBackColor = true;
            // 
            // radioButtonShowLink
            // 
            this.radioButtonShowLink.AutoSize = true;
            this.radioButtonShowLink.Checked = true;
            this.radioButtonShowLink.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.radioButtonShowLink.ForeColor = System.Drawing.Color.Black;
            this.radioButtonShowLink.Location = new System.Drawing.Point(19, 30);
            this.radioButtonShowLink.Name = "radioButtonShowLink";
            this.radioButtonShowLink.Size = new System.Drawing.Size(182, 17);
            this.radioButtonShowLink.TabIndex = 2;
            this.radioButtonShowLink.TabStop = true;
            this.radioButtonShowLink.Text = "Show Online Documentation Link";
            this.radioButtonShowLink.UseVisualStyleBackColor = true;
            this.radioButtonShowLink.CheckedChanged += new System.EventHandler(this.radioButtonShowLink_CheckedChanged);
            // 
            // OptionsForm
            // 
            this.AcceptButton = this.buttonDone;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.CancelButton = this.buttonDone;
            this.ClientSize = new System.Drawing.Size(429, 199);
            this.Controls.Add(this.groupBoxOnlineMode);
            this.Controls.Add(this.buttonDone);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OptionsForm";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Options";
            this.groupBoxOnlineMode.ResumeLayout(false);
            this.groupBoxOnlineMode.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonDone;
        private System.Windows.Forms.GroupBox groupBoxOnlineMode;
        private System.Windows.Forms.RadioButton radioButtonConnect;
        private System.Windows.Forms.RadioButton radioButtonShowLink;

    }
}
