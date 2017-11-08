namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    partial class FinishControl
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FinishControl));
            this.animatedPanel1 = new NetOffice.DeveloperToolbox.Utils.Animation.Panel.AnimatedPanel();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.labelMessage = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.buttonClose = new NetOffice.DeveloperToolbox.Controls.Buttons.RoundedButton();
            this.buttonOpenFolder = new NetOffice.DeveloperToolbox.Controls.Buttons.RoundedButton();
            this.buttonOpenSolution = new NetOffice.DeveloperToolbox.Controls.Buttons.RoundedButton();
            this.controlBackColorAnimator1 = new NetOffice.DeveloperToolbox.Utils.Animation.Colors.ControlBackColorAnimator(this.components);
            this.controlForeColorAnimator1 = new NetOffice.DeveloperToolbox.Utils.Animation.ControlForeColorAnimator(this.components);
            this.animatedPanel1.SuspendLayout();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panelButtons.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.controlBackColorAnimator1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.controlForeColorAnimator1)).BeginInit();
            this.SuspendLayout();
            // 
            // animatedPanel1
            // 
            this.animatedPanel1.Animation1Enabled = true;
            this.animatedPanel1.Animation1Image = ((System.Drawing.Image)(resources.GetObject("animatedPanel1.Animation1Image")));
            this.animatedPanel1.Animation1ImageCount = 6;
            this.animatedPanel1.Animation1Intervall = 50;
            this.animatedPanel1.Animation1Opacity = 0.1F;
            this.animatedPanel1.Animation2Enabled = true;
            this.animatedPanel1.Animation2Image = ((System.Drawing.Image)(resources.GetObject("animatedPanel1.Animation2Image")));
            this.animatedPanel1.Animation2ImageCount = 3;
            this.animatedPanel1.Animation2Opacity = 0.1F;
            this.animatedPanel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.animatedPanel1.Controls.Add(this.panelHeader);
            this.animatedPanel1.Controls.Add(this.panelButtons);
            this.animatedPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.animatedPanel1.Location = new System.Drawing.Point(0, 0);
            this.animatedPanel1.Name = "animatedPanel1";
            this.animatedPanel1.Size = new System.Drawing.Size(744, 279);
            this.animatedPanel1.TabIndex = 3;
            // 
            // panelHeader
            // 
            this.panelHeader.BackColor = System.Drawing.Color.Transparent;
            this.panelHeader.Controls.Add(this.labelMessage);
            this.panelHeader.Controls.Add(this.pictureBox1);
            this.panelHeader.Location = new System.Drawing.Point(161, 26);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(457, 46);
            this.panelHeader.TabIndex = 2;
            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = true;
            this.labelMessage.BackColor = System.Drawing.Color.Transparent;
            this.labelMessage.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMessage.ForeColor = System.Drawing.Color.SteelBlue;
            this.labelMessage.Location = new System.Drawing.Point(65, 11);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(342, 25);
            this.labelMessage.TabIndex = 30;
            this.labelMessage.Text = "The Project is successfully completed.";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(12, 7);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(37, 32);
            this.pictureBox1.TabIndex = 29;
            this.pictureBox1.TabStop = false;
            // 
            // panelButtons
            // 
            this.panelButtons.BackColor = System.Drawing.Color.Transparent;
            this.panelButtons.Controls.Add(this.buttonClose);
            this.panelButtons.Controls.Add(this.buttonOpenFolder);
            this.panelButtons.Controls.Add(this.buttonOpenSolution);
            this.panelButtons.Location = new System.Drawing.Point(241, 108);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Size = new System.Drawing.Size(255, 156);
            this.panelButtons.TabIndex = 1;
            // 
            // buttonClose
            // 
            this.buttonClose.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonClose.FlatAppearance.BorderSize = 2;
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.ForeColor = System.Drawing.Color.Black;
            this.buttonClose.Image = ((System.Drawing.Image)(resources.GetObject("buttonClose.Image")));
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.Location = new System.Drawing.Point(6, 117);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(244, 36);
            this.buttonClose.TabIndex = 2;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // buttonOpenFolder
            // 
            this.buttonOpenFolder.Enabled = false;
            this.buttonOpenFolder.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonOpenFolder.FlatAppearance.BorderSize = 2;
            this.buttonOpenFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOpenFolder.ForeColor = System.Drawing.Color.Black;
            this.buttonOpenFolder.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpenFolder.Image")));
            this.buttonOpenFolder.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonOpenFolder.Location = new System.Drawing.Point(6, 61);
            this.buttonOpenFolder.Name = "buttonOpenFolder";
            this.buttonOpenFolder.Size = new System.Drawing.Size(244, 36);
            this.buttonOpenFolder.TabIndex = 1;
            this.buttonOpenFolder.Text = "Open Folder";
            this.buttonOpenFolder.UseVisualStyleBackColor = true;
            this.buttonOpenFolder.Click += new System.EventHandler(this.buttonOpenFolder_Click);
            // 
            // buttonOpenSolution
            // 
            this.buttonOpenSolution.Enabled = false;
            this.buttonOpenSolution.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonOpenSolution.FlatAppearance.BorderSize = 2;
            this.buttonOpenSolution.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOpenSolution.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonOpenSolution.ForeColor = System.Drawing.Color.Black;
            this.buttonOpenSolution.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpenSolution.Image")));
            this.buttonOpenSolution.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonOpenSolution.Location = new System.Drawing.Point(6, 4);
            this.buttonOpenSolution.Name = "buttonOpenSolution";
            this.buttonOpenSolution.Size = new System.Drawing.Size(244, 36);
            this.buttonOpenSolution.TabIndex = 0;
            this.buttonOpenSolution.Text = "Open Solution";
            this.buttonOpenSolution.UseVisualStyleBackColor = true;
            this.buttonOpenSolution.Click += new System.EventHandler(this.buttonOpenSolution_Click);
            // 
            // controlBackColorAnimator1
            // 
            this.controlBackColorAnimator1.Control = this.buttonOpenSolution;
            this.controlBackColorAnimator1.EndColor = System.Drawing.Color.CornflowerBlue;
            this.controlBackColorAnimator1.Intervall = 30;
            this.controlBackColorAnimator1.LoopMode = NetOffice.DeveloperToolbox.Utils.Animation.LoopMode.Bidirectional;
            this.controlBackColorAnimator1.StartColor = System.Drawing.Color.LightSteelBlue;
            // 
            // controlForeColorAnimator1
            // 
            this.controlForeColorAnimator1.Control = this.labelMessage;
            this.controlForeColorAnimator1.EndColor = System.Drawing.Color.DodgerBlue;
            this.controlForeColorAnimator1.Intervall = 35;
            this.controlForeColorAnimator1.LoopMode = NetOffice.DeveloperToolbox.Utils.Animation.LoopMode.Bidirectional;
            this.controlForeColorAnimator1.StartColor = System.Drawing.Color.SteelBlue;
            this.controlForeColorAnimator1.StepSize = 1D;
            // 
            // FinishControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.animatedPanel1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "FinishControl";
            this.Size = new System.Drawing.Size(744, 279);
            this.Resize += new System.EventHandler(this.FinishControl_Resize);
            this.animatedPanel1.ResumeLayout(false);
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panelButtons.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.controlBackColorAnimator1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.controlForeColorAnimator1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DeveloperToolbox.Controls.Buttons.RoundedButton buttonOpenSolution;
        private System.Windows.Forms.Panel panelButtons;
        private DeveloperToolbox.Controls.Buttons.RoundedButton buttonOpenFolder;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.PictureBox pictureBox1;
        private DeveloperToolbox.Controls.Buttons.RoundedButton buttonClose;
        private Utils.Animation.Panel.AnimatedPanel animatedPanel1;
        private Utils.Animation.Colors.ControlBackColorAnimator controlBackColorAnimator1;
        private Utils.Animation.ControlForeColorAnimator controlForeColorAnimator1;
    }
}
