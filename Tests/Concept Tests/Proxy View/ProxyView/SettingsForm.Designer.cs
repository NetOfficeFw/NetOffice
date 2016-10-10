namespace ProxyView
{
    partial class SettingsForm
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
            this.IntervalTrackBar = new System.Windows.Forms.TrackBar();
            this.DetailsCheckBox = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.IntervalLabel = new System.Windows.Forms.Label();
            this.ApplyButton = new System.Windows.Forms.Button();
            this.DiscardButton = new System.Windows.Forms.Button();
            this.ShowAllAccessibleButton = new System.Windows.Forms.RadioButton();
            this.ShowOfficeAccessibleButton = new System.Windows.Forms.RadioButton();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.IntervalTrackBar)).BeginInit();
            this.SuspendLayout();
            // 
            // IntervalTrackBar
            // 
            this.IntervalTrackBar.Location = new System.Drawing.Point(30, 64);
            this.IntervalTrackBar.Maximum = 90000;
            this.IntervalTrackBar.Minimum = 1000;
            this.IntervalTrackBar.Name = "IntervalTrackBar";
            this.IntervalTrackBar.Size = new System.Drawing.Size(322, 45);
            this.IntervalTrackBar.TabIndex = 0;
            this.IntervalTrackBar.Value = 5000;
            this.IntervalTrackBar.ValueChanged += new System.EventHandler(this.IntervalTrackBar_ValueChanged);
            // 
            // DetailsCheckBox
            // 
            this.DetailsCheckBox.AutoSize = true;
            this.DetailsCheckBox.Location = new System.Drawing.Point(43, 199);
            this.DetailsCheckBox.Name = "DetailsCheckBox";
            this.DetailsCheckBox.Size = new System.Drawing.Size(317, 17);
            this.DetailsCheckBox.TabIndex = 1;
            this.DetailsCheckBox.Text = "Show selected proxy in property grid (performance slow down)";
            this.DetailsCheckBox.UseVisualStyleBackColor = true;
            this.DetailsCheckBox.CheckedChanged += new System.EventHandler(this.DetailsCheckBox_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label1.Location = new System.Drawing.Point(27, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "Refresh Interval";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label2.Location = new System.Drawing.Point(27, 159);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(102, 20);
            this.label2.TabIndex = 3;
            this.label2.Text = "Show Details";
            // 
            // IntervalLabel
            // 
            this.IntervalLabel.AutoSize = true;
            this.IntervalLabel.Location = new System.Drawing.Point(40, 112);
            this.IntervalLabel.Name = "IntervalLabel";
            this.IntervalLabel.Size = new System.Drawing.Size(66, 13);
            this.IntervalLabel.TabIndex = 4;
            this.IntervalLabel.Text = "{0} Seconds";
            // 
            // ApplyButton
            // 
            this.ApplyButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.ApplyButton.Enabled = false;
            this.ApplyButton.Location = new System.Drawing.Point(135, 355);
            this.ApplyButton.Name = "ApplyButton";
            this.ApplyButton.Size = new System.Drawing.Size(95, 29);
            this.ApplyButton.TabIndex = 5;
            this.ApplyButton.Text = "Apply && Close";
            this.ApplyButton.UseVisualStyleBackColor = true;
            this.ApplyButton.Click += new System.EventHandler(this.ApplyButton_Click);
            // 
            // DiscardButton
            // 
            this.DiscardButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.DiscardButton.Location = new System.Drawing.Point(257, 355);
            this.DiscardButton.Name = "DiscardButton";
            this.DiscardButton.Size = new System.Drawing.Size(95, 29);
            this.DiscardButton.TabIndex = 6;
            this.DiscardButton.Text = "Discard && Close";
            this.DiscardButton.UseVisualStyleBackColor = true;
            this.DiscardButton.Click += new System.EventHandler(this.DiscardButton_Click);
            // 
            // ShowAllAccessibleButton
            // 
            this.ShowAllAccessibleButton.AutoSize = true;
            this.ShowAllAccessibleButton.Checked = true;
            this.ShowAllAccessibleButton.Location = new System.Drawing.Point(43, 290);
            this.ShowAllAccessibleButton.Name = "ShowAllAccessibleButton";
            this.ShowAllAccessibleButton.Size = new System.Drawing.Size(66, 17);
            this.ShowAllAccessibleButton.TabIndex = 7;
            this.ShowAllAccessibleButton.TabStop = true;
            this.ShowAllAccessibleButton.Text = "Show All";
            this.ShowAllAccessibleButton.UseVisualStyleBackColor = true;
            this.ShowAllAccessibleButton.CheckedChanged += new System.EventHandler(this.AccessibleButton_CheckedChanged);
            // 
            // ShowOfficeAccessibleButton
            // 
            this.ShowOfficeAccessibleButton.AutoSize = true;
            this.ShowOfficeAccessibleButton.Location = new System.Drawing.Point(158, 290);
            this.ShowOfficeAccessibleButton.Name = "ShowOfficeAccessibleButton";
            this.ShowOfficeAccessibleButton.Size = new System.Drawing.Size(178, 17);
            this.ShowOfficeAccessibleButton.TabIndex = 8;
            this.ShowOfficeAccessibleButton.Text = "Show Supported Office Versions";
            this.ShowOfficeAccessibleButton.UseVisualStyleBackColor = true;
            this.ShowOfficeAccessibleButton.CheckedChanged += new System.EventHandler(this.AccessibleButton_CheckedChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label3.Location = new System.Drawing.Point(27, 244);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(90, 20);
            this.label3.TabIndex = 9;
            this.label3.Text = "IAccessible";
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.ApplyButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.DiscardButton;
            this.ClientSize = new System.Drawing.Size(384, 412);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ShowOfficeAccessibleButton);
            this.Controls.Add(this.ShowAllAccessibleButton);
            this.Controls.Add(this.DiscardButton);
            this.Controls.Add(this.ApplyButton);
            this.Controls.Add(this.IntervalLabel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.DetailsCheckBox);
            this.Controls.Add(this.IntervalTrackBar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Advanced Settings";
            ((System.ComponentModel.ISupportInitialize)(this.IntervalTrackBar)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TrackBar IntervalTrackBar;
        private System.Windows.Forms.CheckBox DetailsCheckBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label IntervalLabel;
        private System.Windows.Forms.Button ApplyButton;
        private System.Windows.Forms.Button DiscardButton;
        private System.Windows.Forms.RadioButton ShowAllAccessibleButton;
        private System.Windows.Forms.RadioButton ShowOfficeAccessibleButton;
        private System.Windows.Forms.Label label3;
    }
}