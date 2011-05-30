namespace NetOffice.DeveloperUtils.SupportByLibrary
{
    partial class SupportByLibraryControl
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SupportByLibraryControl));
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxAssembly = new System.Windows.Forms.TextBox();
            this.buttonSelectAssembly = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.Application = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column09 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Column10 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Column11 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Column12 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Column14 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColumnSet = new System.Windows.Forms.DataGridViewButtonColumn();
            this.ColumnClear = new System.Windows.Forms.DataGridViewButtonColumn();
            this.buttonInfo = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.textBoxConsole = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.textBoxDocument = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Assembly";
            // 
            // textBoxAssembly
            // 
            this.textBoxAssembly.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxAssembly.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.textBoxAssembly.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxAssembly.Location = new System.Drawing.Point(64, 11);
            this.textBoxAssembly.Name = "textBoxAssembly";
            this.textBoxAssembly.ReadOnly = true;
            this.textBoxAssembly.Size = new System.Drawing.Size(289, 20);
            this.textBoxAssembly.TabIndex = 1;
            // 
            // buttonSelectAssembly
            // 
            this.buttonSelectAssembly.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSelectAssembly.Location = new System.Drawing.Point(363, 9);
            this.buttonSelectAssembly.Name = "buttonSelectAssembly";
            this.buttonSelectAssembly.Size = new System.Drawing.Size(55, 23);
            this.buttonSelectAssembly.TabIndex = 2;
            this.buttonSelectAssembly.Text = "Select";
            this.buttonSelectAssembly.UseVisualStyleBackColor = true;
            this.buttonSelectAssembly.Click += new System.EventHandler(this.buttonSelectAssembly_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView.BackgroundColor = System.Drawing.SystemColors.ActiveBorder;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Application,
            this.Column09,
            this.Column10,
            this.Column11,
            this.Column12,
            this.Column14,
            this.ColumnSet,
            this.ColumnClear});
            this.dataGridView.Location = new System.Drawing.Point(8, 47);
            this.dataGridView.MultiSelect = false;
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersVisible = false;
            this.dataGridView.ShowCellErrors = false;
            this.dataGridView.ShowCellToolTips = false;
            this.dataGridView.ShowEditingIcon = false;
            this.dataGridView.ShowRowErrors = false;
            this.dataGridView.Size = new System.Drawing.Size(460, 133);
            this.dataGridView.TabIndex = 4;
            this.dataGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellClick);
            // 
            // Application
            // 
            this.Application.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.Application.DefaultCellStyle = dataGridViewCellStyle7;
            this.Application.HeaderText = "";
            this.Application.Name = "Application";
            this.Application.ReadOnly = true;
            // 
            // Column09
            // 
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.ActiveBorder;
            dataGridViewCellStyle8.NullValue = false;
            this.Column09.DefaultCellStyle = dataGridViewCellStyle8;
            this.Column09.HeaderText = "   09";
            this.Column09.Name = "Column09";
            this.Column09.Width = 50;
            // 
            // Column10
            // 
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.ActiveBorder;
            dataGridViewCellStyle9.NullValue = false;
            this.Column10.DefaultCellStyle = dataGridViewCellStyle9;
            this.Column10.HeaderText = "   10";
            this.Column10.Name = "Column10";
            this.Column10.Width = 50;
            // 
            // Column11
            // 
            dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.ActiveBorder;
            dataGridViewCellStyle10.NullValue = false;
            this.Column11.DefaultCellStyle = dataGridViewCellStyle10;
            this.Column11.HeaderText = "   11";
            this.Column11.Name = "Column11";
            this.Column11.Width = 50;
            // 
            // Column12
            // 
            dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.ActiveBorder;
            dataGridViewCellStyle11.NullValue = false;
            this.Column12.DefaultCellStyle = dataGridViewCellStyle11;
            this.Column12.HeaderText = "   12";
            this.Column12.Name = "Column12";
            this.Column12.Width = 50;
            // 
            // Column14
            // 
            dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.ActiveBorder;
            dataGridViewCellStyle12.NullValue = false;
            this.Column14.DefaultCellStyle = dataGridViewCellStyle12;
            this.Column14.HeaderText = "   14";
            this.Column14.Name = "Column14";
            this.Column14.Width = 50;
            // 
            // ColumnSet
            // 
            this.ColumnSet.HeaderText = "Set All";
            this.ColumnSet.Name = "ColumnSet";
            this.ColumnSet.Width = 50;
            // 
            // ColumnClear
            // 
            this.ColumnClear.HeaderText = "Clear";
            this.ColumnClear.Name = "ColumnClear";
            this.ColumnClear.Width = 50;
            // 
            // buttonInfo
            // 
            this.buttonInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonInfo.Image = ((System.Drawing.Image)(resources.GetObject("buttonInfo.Image")));
            this.buttonInfo.Location = new System.Drawing.Point(440, 5);
            this.buttonInfo.Name = "buttonInfo";
            this.buttonInfo.Size = new System.Drawing.Size(28, 28);
            this.buttonInfo.TabIndex = 27;
            this.buttonInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonInfo.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(8, 186);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(460, 125);
            this.tabControl1.TabIndex = 29;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.textBoxConsole);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(452, 99);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // textBoxConsole
            // 
            this.textBoxConsole.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxConsole.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxConsole.Location = new System.Drawing.Point(3, 3);
            this.textBoxConsole.Multiline = true;
            this.textBoxConsole.Name = "textBoxConsole";
            this.textBoxConsole.ReadOnly = true;
            this.textBoxConsole.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxConsole.Size = new System.Drawing.Size(446, 93);
            this.textBoxConsole.TabIndex = 29;
            this.textBoxConsole.WordWrap = false;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.textBoxDocument);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(452, 99);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // textBoxDocument
            // 
            this.textBoxDocument.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxDocument.Location = new System.Drawing.Point(3, 3);
            this.textBoxDocument.Multiline = true;
            this.textBoxDocument.Name = "textBoxDocument";
            this.textBoxDocument.ReadOnly = true;
            this.textBoxDocument.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxDocument.Size = new System.Drawing.Size(446, 93);
            this.textBoxDocument.TabIndex = 30;
            // 
            // SupportByLibraryControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.buttonInfo);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.buttonSelectAssembly);
            this.Controls.Add(this.textBoxAssembly);
            this.Controls.Add(this.label1);
            this.Name = "SupportByLibraryControl";
            this.Size = new System.Drawing.Size(476, 311);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxAssembly;
        private System.Windows.Forms.Button buttonSelectAssembly;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Button buttonInfo;
        private System.Windows.Forms.DataGridViewTextBoxColumn Application;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column09;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column10;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column11;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column12;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column14;
        private System.Windows.Forms.DataGridViewButtonColumn ColumnSet;
        private System.Windows.Forms.DataGridViewButtonColumn ColumnClear;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TextBox textBoxConsole;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TextBox textBoxDocument;
    }
}
