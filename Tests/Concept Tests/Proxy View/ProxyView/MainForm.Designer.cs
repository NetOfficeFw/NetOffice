namespace ProxyView
{
    partial class MainForm
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.EntriesGrid = new System.Windows.Forms.DataGridView();
            this.GridContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.applicationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.RefreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ExitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AutoRefreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ShowAccessibleInsteadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AdvancedToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.InfoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.EntryDetailsGrid = new System.Windows.Forms.PropertyGrid();
            this.RefreshTimer = new System.Windows.Forms.Timer(this.components);
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EntriesGrid)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.EntriesGrid);
            this.splitContainer1.Panel1.Controls.Add(this.menuStrip1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.EntryDetailsGrid);
            this.splitContainer1.Size = new System.Drawing.Size(589, 460);
            this.splitContainer1.SplitterDistance = 415;
            this.splitContainer1.TabIndex = 0;
            // 
            // EntriesGrid
            // 
            this.EntriesGrid.AllowUserToAddRows = false;
            this.EntriesGrid.AllowUserToDeleteRows = false;
            this.EntriesGrid.AllowUserToOrderColumns = true;
            this.EntriesGrid.AllowUserToResizeRows = false;
            this.EntriesGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.EntriesGrid.BackgroundColor = System.Drawing.Color.LightSteelBlue;
            this.EntriesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EntriesGrid.ContextMenuStrip = this.GridContextMenu;
            this.EntriesGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.EntriesGrid.Location = new System.Drawing.Point(0, 24);
            this.EntriesGrid.MultiSelect = false;
            this.EntriesGrid.Name = "EntriesGrid";
            this.EntriesGrid.RowHeadersVisible = false;
            this.EntriesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.EntriesGrid.Size = new System.Drawing.Size(415, 436);
            this.EntriesGrid.TabIndex = 0;
            this.EntriesGrid.SelectionChanged += new System.EventHandler(this.EntriesGrid_SelectionChanged);
            // 
            // GridContextMenu
            // 
            this.GridContextMenu.Name = "contextMenuStrip1";
            this.GridContextMenu.Size = new System.Drawing.Size(61, 4);
            this.GridContextMenu.Opening += new System.ComponentModel.CancelEventHandler(this.GridContextMenu_Opening);
            this.GridContextMenu.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.GridContextMenu_ItemClicked);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.applicationToolStripMenuItem,
            this.settingsToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(415, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // applicationToolStripMenuItem
            // 
            this.applicationToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshToolStripMenuItem,
            this.toolStripSeparator1,
            this.ExitToolStripMenuItem});
            this.applicationToolStripMenuItem.Name = "applicationToolStripMenuItem";
            this.applicationToolStripMenuItem.Size = new System.Drawing.Size(80, 20);
            this.applicationToolStripMenuItem.Text = "Application";
            // 
            // RefreshToolStripMenuItem
            // 
            this.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem";
            this.RefreshToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.RefreshToolStripMenuItem.Text = "Refresh";
            this.RefreshToolStripMenuItem.Click += new System.EventHandler(this.RefreshToolStripMenuItem_Click);
            // 
            // ExitToolStripMenuItem
            // 
            this.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem";
            this.ExitToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.ExitToolStripMenuItem.Text = "Close";
            this.ExitToolStripMenuItem.Click += new System.EventHandler(this.ExitToolStripMenuItem_Click);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AutoRefreshToolStripMenuItem,
            this.ShowAccessibleInsteadToolStripMenuItem,
            this.AdvancedToolStripMenuItem});
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.settingsToolStripMenuItem.Text = "Settings";
            // 
            // AutoRefreshToolStripMenuItem
            // 
            this.AutoRefreshToolStripMenuItem.CheckOnClick = true;
            this.AutoRefreshToolStripMenuItem.Name = "AutoRefreshToolStripMenuItem";
            this.AutoRefreshToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.AutoRefreshToolStripMenuItem.Text = "Enable Auto Refresh";
            this.AutoRefreshToolStripMenuItem.CheckedChanged += new System.EventHandler(this.AutoRefreshToolStripMenuItem_CheckedChanged);
            // 
            // ShowAccessibleInsteadToolStripMenuItem
            // 
            this.ShowAccessibleInsteadToolStripMenuItem.CheckOnClick = true;
            this.ShowAccessibleInsteadToolStripMenuItem.Name = "ShowAccessibleInsteadToolStripMenuItem";
            this.ShowAccessibleInsteadToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.ShowAccessibleInsteadToolStripMenuItem.Text = "Show Accessible Instead";
            this.ShowAccessibleInsteadToolStripMenuItem.CheckedChanged += new System.EventHandler(this.ShowAccessibleInsteadToolStripMenuItem_CheckedChanged);
            // 
            // AdvancedToolStripMenuItem
            // 
            this.AdvancedToolStripMenuItem.Name = "AdvancedToolStripMenuItem";
            this.AdvancedToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.AdvancedToolStripMenuItem.Text = "Advanced Settings";
            this.AdvancedToolStripMenuItem.Click += new System.EventHandler(this.AdvancedToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.InfoToolStripMenuItem});
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
            this.aboutToolStripMenuItem.Text = "About";
            // 
            // InfoToolStripMenuItem
            // 
            this.InfoToolStripMenuItem.Name = "InfoToolStripMenuItem";
            this.InfoToolStripMenuItem.Size = new System.Drawing.Size(95, 22);
            this.InfoToolStripMenuItem.Text = "Info";
            this.InfoToolStripMenuItem.Click += new System.EventHandler(this.InfoToolStripMenuItem_Click);
            // 
            // EntryDetailsGrid
            // 
            this.EntryDetailsGrid.CategoryForeColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.EntryDetailsGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.EntryDetailsGrid.Location = new System.Drawing.Point(0, 0);
            this.EntryDetailsGrid.Name = "EntryDetailsGrid";
            this.EntryDetailsGrid.PropertySort = System.Windows.Forms.PropertySort.Alphabetical;
            this.EntryDetailsGrid.Size = new System.Drawing.Size(170, 460);
            this.EntryDetailsGrid.TabIndex = 0;
            // 
            // RefreshTimer
            // 
            this.RefreshTimer.Interval = 5000;
            this.RefreshTimer.Tick += new System.EventHandler(this.RefreshTimer_Tick);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(149, 6);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(589, 460);
            this.Controls.Add(this.splitContainer1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.EntriesGrid)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView EntriesGrid;
        private System.Windows.Forms.PropertyGrid EntryDetailsGrid;
        private System.Windows.Forms.Timer RefreshTimer;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem applicationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ExitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AutoRefreshToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AdvancedToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem InfoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem RefreshToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ShowAccessibleInsteadToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip GridContextMenu;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
    }
}

