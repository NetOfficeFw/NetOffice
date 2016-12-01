namespace NetOffice.OfficeApi.Tools.Utils
{
    partial class TrayMenuMonitorItemControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.HighlightTimer = new System.Windows.Forms.Timer(this.components);
            this.ContentPanel = new System.Windows.Forms.Panel();
            this.OptionsGrid = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.ShowKindColumnCheckBox = new System.Windows.Forms.CheckBox();
            this.ShowTimeColumnCheckBox = new System.Windows.Forms.CheckBox();
            this.HighlightHelpLabel = new System.Windows.Forms.Label();
            this.AutoExpandHelpLabel = new System.Windows.Forms.Label();
            this.HighlightCheckBox = new System.Windows.Forms.CheckBox();
            this.AutoExpandCheckBox = new System.Windows.Forms.CheckBox();
            this.HierarchicalGrid = new System.Windows.Forms.TreeView();
            this.OverlayPanel = new System.Windows.Forms.Panel();
            this.OverlayTextBox = new System.Windows.Forms.TextBox();
            this.CloseOverlayButton = new System.Windows.Forms.Button();
            this.EnumeratorGrid = new System.Windows.Forms.DataGridView();
            this.SingleGrid = new System.Windows.Forms.PropertyGrid();
            this.CoreRadioButton = new System.Windows.Forms.RadioButton();
            this.SettingsRadioButton = new System.Windows.Forms.RadioButton();
            this.ConsoleRadioButton = new System.Windows.Forms.RadioButton();
            this.DiagnosticsRadioButton = new System.Windows.Forms.RadioButton();
            this.ProxiesRadioButton = new System.Windows.Forms.RadioButton();
            this.OptionsRadioButton = new System.Windows.Forms.RadioButton();
            this.ContentPanel.SuspendLayout();
            this.OptionsGrid.SuspendLayout();
            this.OverlayPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EnumeratorGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // HighlightTimer
            // 
            this.HighlightTimer.Interval = 90;
            this.HighlightTimer.Tick += new System.EventHandler(this.HighlightTimer_Tick);
            // 
            // ContentPanel
            // 
            this.ContentPanel.BackColor = System.Drawing.Color.Transparent;
            this.ContentPanel.Controls.Add(this.OptionsGrid);
            this.ContentPanel.Controls.Add(this.HierarchicalGrid);
            this.ContentPanel.Controls.Add(this.OverlayPanel);
            this.ContentPanel.Controls.Add(this.EnumeratorGrid);
            this.ContentPanel.Controls.Add(this.SingleGrid);
            this.ContentPanel.Location = new System.Drawing.Point(0, 33);
            this.ContentPanel.Name = "ContentPanel";
            this.ContentPanel.Size = new System.Drawing.Size(485, 287);
            this.ContentPanel.TabIndex = 1;
            // 
            // OptionsGrid
            // 
            this.OptionsGrid.Controls.Add(this.label1);
            this.OptionsGrid.Controls.Add(this.label2);
            this.OptionsGrid.Controls.Add(this.ShowKindColumnCheckBox);
            this.OptionsGrid.Controls.Add(this.ShowTimeColumnCheckBox);
            this.OptionsGrid.Controls.Add(this.HighlightHelpLabel);
            this.OptionsGrid.Controls.Add(this.AutoExpandHelpLabel);
            this.OptionsGrid.Controls.Add(this.HighlightCheckBox);
            this.OptionsGrid.Controls.Add(this.AutoExpandCheckBox);
            this.OptionsGrid.Location = new System.Drawing.Point(48, 4);
            this.OptionsGrid.Name = "OptionsGrid";
            this.OptionsGrid.Size = new System.Drawing.Size(434, 206);
            this.OptionsGrid.TabIndex = 4;
            this.OptionsGrid.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label1.Location = new System.Drawing.Point(254, 120);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(129, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Message Kind Information";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label2.Location = new System.Drawing.Point(255, 84);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Message Time Information";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // ShowKindColumnCheckBox
            // 
            this.ShowKindColumnCheckBox.AutoSize = true;
            this.ShowKindColumnCheckBox.Location = new System.Drawing.Point(24, 120);
            this.ShowKindColumnCheckBox.Name = "ShowKindColumnCheckBox";
            this.ShowKindColumnCheckBox.Size = new System.Drawing.Size(156, 17);
            this.ShowKindColumnCheckBox.TabIndex = 7;
            this.ShowKindColumnCheckBox.Text = "Show Console Kind Column";
            this.ShowKindColumnCheckBox.UseVisualStyleBackColor = true;
            // 
            // ShowTimeColumnCheckBox
            // 
            this.ShowTimeColumnCheckBox.AutoSize = true;
            this.ShowTimeColumnCheckBox.Location = new System.Drawing.Point(24, 84);
            this.ShowTimeColumnCheckBox.Name = "ShowTimeColumnCheckBox";
            this.ShowTimeColumnCheckBox.Size = new System.Drawing.Size(158, 17);
            this.ShowTimeColumnCheckBox.TabIndex = 6;
            this.ShowTimeColumnCheckBox.Text = "Show Console Time Column";
            this.ShowTimeColumnCheckBox.UseVisualStyleBackColor = true;
            // 
            // HighlightHelpLabel
            // 
            this.HighlightHelpLabel.AutoSize = true;
            this.HighlightHelpLabel.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.HighlightHelpLabel.Location = new System.Drawing.Point(255, 48);
            this.HighlightHelpLabel.Name = "HighlightHelpLabel";
            this.HighlightHelpLabel.Size = new System.Drawing.Size(110, 13);
            this.HighlightHelpLabel.TabIndex = 5;
            this.HighlightHelpLabel.Text = "Highlight New Proxies";
            this.HighlightHelpLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // AutoExpandHelpLabel
            // 
            this.AutoExpandHelpLabel.AutoSize = true;
            this.AutoExpandHelpLabel.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.AutoExpandHelpLabel.Location = new System.Drawing.Point(255, 15);
            this.AutoExpandHelpLabel.Name = "AutoExpandHelpLabel";
            this.AutoExpandHelpLabel.Size = new System.Drawing.Size(105, 13);
            this.AutoExpandHelpLabel.TabIndex = 4;
            this.AutoExpandHelpLabel.Text = "Auto Expand Proxies";
            this.AutoExpandHelpLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // HighlightCheckBox
            // 
            this.HighlightCheckBox.AutoSize = true;
            this.HighlightCheckBox.Checked = true;
            this.HighlightCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.HighlightCheckBox.Location = new System.Drawing.Point(24, 48);
            this.HighlightCheckBox.Name = "HighlightCheckBox";
            this.HighlightCheckBox.Size = new System.Drawing.Size(67, 17);
            this.HighlightCheckBox.TabIndex = 2;
            this.HighlightCheckBox.Text = "Highlight";
            this.HighlightCheckBox.UseVisualStyleBackColor = true;
            // 
            // AutoExpandCheckBox
            // 
            this.AutoExpandCheckBox.AutoSize = true;
            this.AutoExpandCheckBox.Checked = true;
            this.AutoExpandCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AutoExpandCheckBox.Location = new System.Drawing.Point(24, 15);
            this.AutoExpandCheckBox.Name = "AutoExpandCheckBox";
            this.AutoExpandCheckBox.Size = new System.Drawing.Size(87, 17);
            this.AutoExpandCheckBox.TabIndex = 1;
            this.AutoExpandCheckBox.Text = "Auto Expand";
            this.AutoExpandCheckBox.UseVisualStyleBackColor = true;
            // 
            // HierarchicalGrid
            // 
            this.HierarchicalGrid.BackColor = System.Drawing.Color.White;
            this.HierarchicalGrid.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.HierarchicalGrid.Location = new System.Drawing.Point(34, 1);
            this.HierarchicalGrid.Margin = new System.Windows.Forms.Padding(0);
            this.HierarchicalGrid.Name = "HierarchicalGrid";
            this.HierarchicalGrid.Size = new System.Drawing.Size(11, 110);
            this.HierarchicalGrid.TabIndex = 3;
            this.HierarchicalGrid.Visible = false;
            // 
            // OverlayPanel
            // 
            this.OverlayPanel.Controls.Add(this.OverlayTextBox);
            this.OverlayPanel.Controls.Add(this.CloseOverlayButton);
            this.OverlayPanel.Location = new System.Drawing.Point(2, 213);
            this.OverlayPanel.Margin = new System.Windows.Forms.Padding(0);
            this.OverlayPanel.Name = "OverlayPanel";
            this.OverlayPanel.Size = new System.Drawing.Size(483, 72);
            this.OverlayPanel.TabIndex = 2;
            this.OverlayPanel.Visible = false;
            // 
            // OverlayTextBox
            // 
            this.OverlayTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.OverlayTextBox.BackColor = System.Drawing.Color.White;
            this.OverlayTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.OverlayTextBox.Location = new System.Drawing.Point(0, 0);
            this.OverlayTextBox.Margin = new System.Windows.Forms.Padding(0);
            this.OverlayTextBox.Multiline = true;
            this.OverlayTextBox.Name = "OverlayTextBox";
            this.OverlayTextBox.ReadOnly = true;
            this.OverlayTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.OverlayTextBox.Size = new System.Drawing.Size(481, 21);
            this.OverlayTextBox.TabIndex = 1;
            // 
            // CloseOverlayButton
            // 
            this.CloseOverlayButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.CloseOverlayButton.Location = new System.Drawing.Point(205, 31);
            this.CloseOverlayButton.Name = "CloseOverlayButton";
            this.CloseOverlayButton.Size = new System.Drawing.Size(79, 29);
            this.CloseOverlayButton.TabIndex = 0;
            this.CloseOverlayButton.Text = "Close";
            this.CloseOverlayButton.UseVisualStyleBackColor = true;
            this.CloseOverlayButton.Click += new System.EventHandler(this.CloseOverlayButton_Click);
            // 
            // EnumeratorGrid
            // 
            this.EnumeratorGrid.AllowUserToAddRows = false;
            this.EnumeratorGrid.AllowUserToDeleteRows = false;
            this.EnumeratorGrid.AllowUserToResizeRows = false;
            this.EnumeratorGrid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.EnumeratorGrid.BackgroundColor = System.Drawing.Color.White;
            this.EnumeratorGrid.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.EnumeratorGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EnumeratorGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.EnumeratorGrid.Location = new System.Drawing.Point(19, 1);
            this.EnumeratorGrid.Margin = new System.Windows.Forms.Padding(0);
            this.EnumeratorGrid.MultiSelect = false;
            this.EnumeratorGrid.Name = "EnumeratorGrid";
            this.EnumeratorGrid.RowHeadersVisible = false;
            this.EnumeratorGrid.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.EnumeratorGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.EnumeratorGrid.Size = new System.Drawing.Size(13, 110);
            this.EnumeratorGrid.TabIndex = 1;
            this.EnumeratorGrid.Visible = false;
            this.EnumeratorGrid.DoubleClick += new System.EventHandler(this.EnumeratorGrid_DoubleClick);
            // 
            // SingleGrid
            // 
            this.SingleGrid.CategoryForeColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.SingleGrid.Location = new System.Drawing.Point(3, 1);
            this.SingleGrid.Name = "SingleGrid";
            this.SingleGrid.PropertySort = System.Windows.Forms.PropertySort.Alphabetical;
            this.SingleGrid.Size = new System.Drawing.Size(11, 110);
            this.SingleGrid.TabIndex = 0;
            this.SingleGrid.ToolbarVisible = false;
            // 
            // CoreRadioButton
            // 
            this.CoreRadioButton.Appearance = System.Windows.Forms.Appearance.Button;
            this.CoreRadioButton.Checked = true;
            this.CoreRadioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.CoreRadioButton.Location = new System.Drawing.Point(0, 0);
            this.CoreRadioButton.Name = "CoreRadioButton";
            this.CoreRadioButton.Size = new System.Drawing.Size(80, 30);
            this.CoreRadioButton.TabIndex = 2;
            this.CoreRadioButton.TabStop = true;
            this.CoreRadioButton.Text = "Core";
            this.CoreRadioButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.CoreRadioButton.UseVisualStyleBackColor = true;
            this.CoreRadioButton.CheckedChanged += new System.EventHandler(this.HeaderRadioButton_CheckedChanged);
            // 
            // SettingsRadioButton
            // 
            this.SettingsRadioButton.Appearance = System.Windows.Forms.Appearance.Button;
            this.SettingsRadioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SettingsRadioButton.Location = new System.Drawing.Point(81, 0);
            this.SettingsRadioButton.Name = "SettingsRadioButton";
            this.SettingsRadioButton.Size = new System.Drawing.Size(80, 30);
            this.SettingsRadioButton.TabIndex = 3;
            this.SettingsRadioButton.Text = "Settings";
            this.SettingsRadioButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.SettingsRadioButton.UseVisualStyleBackColor = true;
            this.SettingsRadioButton.CheckedChanged += new System.EventHandler(this.HeaderRadioButton_CheckedChanged);
            // 
            // ConsoleRadioButton
            // 
            this.ConsoleRadioButton.Appearance = System.Windows.Forms.Appearance.Button;
            this.ConsoleRadioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ConsoleRadioButton.Location = new System.Drawing.Point(162, 0);
            this.ConsoleRadioButton.Name = "ConsoleRadioButton";
            this.ConsoleRadioButton.Size = new System.Drawing.Size(80, 30);
            this.ConsoleRadioButton.TabIndex = 4;
            this.ConsoleRadioButton.Text = "Console";
            this.ConsoleRadioButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.ConsoleRadioButton.UseVisualStyleBackColor = true;
            this.ConsoleRadioButton.CheckedChanged += new System.EventHandler(this.HeaderRadioButton_CheckedChanged);
            // 
            // DiagnosticsRadioButton
            // 
            this.DiagnosticsRadioButton.Appearance = System.Windows.Forms.Appearance.Button;
            this.DiagnosticsRadioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DiagnosticsRadioButton.Location = new System.Drawing.Point(243, 0);
            this.DiagnosticsRadioButton.Name = "DiagnosticsRadioButton";
            this.DiagnosticsRadioButton.Size = new System.Drawing.Size(80, 30);
            this.DiagnosticsRadioButton.TabIndex = 5;
            this.DiagnosticsRadioButton.Text = "Diags";
            this.DiagnosticsRadioButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.DiagnosticsRadioButton.UseVisualStyleBackColor = true;
            this.DiagnosticsRadioButton.CheckedChanged += new System.EventHandler(this.HeaderRadioButton_CheckedChanged);
            // 
            // ProxiesRadioButton
            // 
            this.ProxiesRadioButton.Appearance = System.Windows.Forms.Appearance.Button;
            this.ProxiesRadioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ProxiesRadioButton.Location = new System.Drawing.Point(324, 0);
            this.ProxiesRadioButton.Name = "ProxiesRadioButton";
            this.ProxiesRadioButton.Size = new System.Drawing.Size(80, 30);
            this.ProxiesRadioButton.TabIndex = 6;
            this.ProxiesRadioButton.Text = "Proxies";
            this.ProxiesRadioButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.ProxiesRadioButton.UseVisualStyleBackColor = true;
            this.ProxiesRadioButton.CheckedChanged += new System.EventHandler(this.HeaderRadioButton_CheckedChanged);
            // 
            // OptionsRadioButton
            // 
            this.OptionsRadioButton.Appearance = System.Windows.Forms.Appearance.Button;
            this.OptionsRadioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OptionsRadioButton.Location = new System.Drawing.Point(405, 0);
            this.OptionsRadioButton.Name = "OptionsRadioButton";
            this.OptionsRadioButton.Size = new System.Drawing.Size(80, 30);
            this.OptionsRadioButton.TabIndex = 7;
            this.OptionsRadioButton.Text = "Options";
            this.OptionsRadioButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.OptionsRadioButton.UseVisualStyleBackColor = true;
            this.OptionsRadioButton.CheckedChanged += new System.EventHandler(this.HeaderRadioButton_CheckedChanged);
            // 
            // TrayMenuMonitorItemControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.OptionsRadioButton);
            this.Controls.Add(this.ProxiesRadioButton);
            this.Controls.Add(this.DiagnosticsRadioButton);
            this.Controls.Add(this.ConsoleRadioButton);
            this.Controls.Add(this.SettingsRadioButton);
            this.Controls.Add(this.CoreRadioButton);
            this.Controls.Add(this.ContentPanel);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "TrayMenuMonitorItemControl";
            this.Size = new System.Drawing.Size(485, 320);
            this.ContentPanel.ResumeLayout(false);
            this.OptionsGrid.ResumeLayout(false);
            this.OptionsGrid.PerformLayout();
            this.OverlayPanel.ResumeLayout(false);
            this.OverlayPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EnumeratorGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Timer HighlightTimer;
        private System.Windows.Forms.Panel ContentPanel;
        private System.Windows.Forms.RadioButton CoreRadioButton;
        private System.Windows.Forms.RadioButton SettingsRadioButton;
        private System.Windows.Forms.RadioButton ConsoleRadioButton;
        private System.Windows.Forms.RadioButton DiagnosticsRadioButton;
        private System.Windows.Forms.DataGridView EnumeratorGrid;
        private System.Windows.Forms.PropertyGrid SingleGrid;
        private System.Windows.Forms.Panel OverlayPanel;
        private System.Windows.Forms.TextBox OverlayTextBox;
        private System.Windows.Forms.Button CloseOverlayButton;
        private System.Windows.Forms.RadioButton ProxiesRadioButton;
        private System.Windows.Forms.RadioButton OptionsRadioButton;
        private System.Windows.Forms.TreeView HierarchicalGrid;
        private System.Windows.Forms.Panel OptionsGrid;
        private System.Windows.Forms.CheckBox AutoExpandCheckBox;
        private System.Windows.Forms.CheckBox HighlightCheckBox;
        private System.Windows.Forms.Label AutoExpandHelpLabel;
        private System.Windows.Forms.Label HighlightHelpLabel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox ShowKindColumnCheckBox;
        private System.Windows.Forms.CheckBox ShowTimeColumnCheckBox;
    }
}
