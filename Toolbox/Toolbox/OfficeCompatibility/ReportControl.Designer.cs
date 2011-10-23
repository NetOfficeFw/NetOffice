namespace NetOffice.DeveloperToolbox.OfficeCompatibility
{
    partial class ReportControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportControl));
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.treeViewReport = new System.Windows.Forms.TreeView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.listView1 = new System.Windows.Forms.ListView();
            this.textBoxReport = new System.Windows.Forms.TextBox();
            this.buttonClose2 = new System.Windows.Forms.Button();
            this.labelFilterCaption = new System.Windows.Forms.Label();
            this.comboBoxFilter = new System.Windows.Forms.ComboBox();
            this.pictureBox8 = new System.Windows.Forms.PictureBox();
            this.labelFilterHint = new System.Windows.Forms.Label();
            this.panelView = new System.Windows.Forms.Panel();
            this.panelNativeView = new System.Windows.Forms.Panel();
            this.checkBoxNativeView = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox8)).BeginInit();
            this.panelView.SuspendLayout();
            this.panelNativeView.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(1, 44);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.treeViewReport);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.panelNativeView);
            this.splitContainer1.Panel2.Controls.Add(this.panelView);
            this.splitContainer1.Size = new System.Drawing.Size(797, 385);
            this.splitContainer1.SplitterDistance = 265;
            this.splitContainer1.TabIndex = 0;
            // 
            // treeViewReport
            // 
            this.treeViewReport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewReport.ImageIndex = 0;
            this.treeViewReport.ImageList = this.imageList1;
            this.treeViewReport.Location = new System.Drawing.Point(0, 0);
            this.treeViewReport.Name = "treeViewReport";
            this.treeViewReport.SelectedImageIndex = 0;
            this.treeViewReport.Size = new System.Drawing.Size(265, 385);
            this.treeViewReport.TabIndex = 0;
            this.treeViewReport.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeViewReport_AfterSelect);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Assembly.ICO");
            this.imageList1.Images.SetKeyName(1, "Module.png");
            this.imageList1.Images.SetKeyName(2, "Field.png");
            this.imageList1.Images.SetKeyName(3, "Method.png");
            // 
            // listView1
            // 
            this.listView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listView1.FullRowSelect = true;
            this.listView1.Location = new System.Drawing.Point(0, 0);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(497, 168);
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // textBoxReport
            // 
            this.textBoxReport.BackColor = System.Drawing.Color.White;
            this.textBoxReport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxReport.Location = new System.Drawing.Point(0, 0);
            this.textBoxReport.Multiline = true;
            this.textBoxReport.Name = "textBoxReport";
            this.textBoxReport.ReadOnly = true;
            this.textBoxReport.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxReport.Size = new System.Drawing.Size(497, 93);
            this.textBoxReport.TabIndex = 0;
            this.textBoxReport.WordWrap = false;
            // 
            // buttonClose2
            // 
            this.buttonClose2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonClose2.Location = new System.Drawing.Point(753, 10);
            this.buttonClose2.Name = "buttonClose2";
            this.buttonClose2.Size = new System.Drawing.Size(28, 28);
            this.buttonClose2.TabIndex = 30;
            this.buttonClose2.Text = "X";
            this.buttonClose2.UseVisualStyleBackColor = true;
            this.buttonClose2.Click += new System.EventHandler(this.buttonClose2_Click);
            // 
            // labelFilterCaption
            // 
            this.labelFilterCaption.AutoSize = true;
            this.labelFilterCaption.Location = new System.Drawing.Point(235, 16);
            this.labelFilterCaption.Name = "labelFilterCaption";
            this.labelFilterCaption.Size = new System.Drawing.Size(72, 13);
            this.labelFilterCaption.TabIndex = 31;
            this.labelFilterCaption.Text = "Negativ-Filter:";
            // 
            // comboBoxFilter
            // 
            this.comboBoxFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxFilter.FormattingEnabled = true;
            this.comboBoxFilter.Items.AddRange(new object[] {
            "No Filter",
            "Office 2000 (09)",
            "Office 2002 (10)",
            "Office 2003 (11)",
            "Office 2007 (12)",
            "Office 2010 (14)"});
            this.comboBoxFilter.Location = new System.Drawing.Point(313, 10);
            this.comboBoxFilter.Name = "comboBoxFilter";
            this.comboBoxFilter.Size = new System.Drawing.Size(123, 21);
            this.comboBoxFilter.TabIndex = 32;
            this.comboBoxFilter.SelectedIndexChanged += new System.EventHandler(this.comboBoxFilter_SelectedIndexChanged);
            // 
            // pictureBox8
            // 
            this.pictureBox8.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox8.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox8.Image")));
            this.pictureBox8.Location = new System.Drawing.Point(446, 9);
            this.pictureBox8.Name = "pictureBox8";
            this.pictureBox8.Size = new System.Drawing.Size(16, 16);
            this.pictureBox8.TabIndex = 72;
            this.pictureBox8.TabStop = false;
            // 
            // labelFilterHint
            // 
            this.labelFilterHint.BackColor = System.Drawing.Color.Khaki;
            this.labelFilterHint.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelFilterHint.Location = new System.Drawing.Point(465, 10);
            this.labelFilterHint.Name = "labelFilterHint";
            this.labelFilterHint.Size = new System.Drawing.Size(282, 26);
            this.labelFilterHint.TabIndex = 71;
            this.labelFilterHint.Text = "Nutzen Sie den Filter um nur Funktionalitäten anzuzeigen die eine bestimmte Versi" +
                "on nicht unterstützen.";
            // 
            // panelView
            // 
            this.panelView.Controls.Add(this.listView1);
            this.panelView.Location = new System.Drawing.Point(28, 12);
            this.panelView.Name = "panelView";
            this.panelView.Size = new System.Drawing.Size(497, 168);
            this.panelView.TabIndex = 0;
            // 
            // panelNativeView
            // 
            this.panelNativeView.Controls.Add(this.textBoxReport);
            this.panelNativeView.Location = new System.Drawing.Point(28, 229);
            this.panelNativeView.Name = "panelNativeView";
            this.panelNativeView.Size = new System.Drawing.Size(497, 93);
            this.panelNativeView.TabIndex = 73;
            this.panelNativeView.Visible = false;
            // 
            // checkBoxNativeView
            // 
            this.checkBoxNativeView.AutoSize = true;
            this.checkBoxNativeView.Location = new System.Drawing.Point(23, 12);
            this.checkBoxNativeView.Name = "checkBoxNativeView";
            this.checkBoxNativeView.Size = new System.Drawing.Size(95, 17);
            this.checkBoxNativeView.TabIndex = 73;
            this.checkBoxNativeView.Text = "Native Ansicht";
            this.checkBoxNativeView.UseVisualStyleBackColor = true;
            this.checkBoxNativeView.CheckedChanged += new System.EventHandler(this.checkBoxNativeView_CheckedChanged);
            // 
            // ReportControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.checkBoxNativeView);
            this.Controls.Add(this.pictureBox8);
            this.Controls.Add(this.labelFilterHint);
            this.Controls.Add(this.comboBoxFilter);
            this.Controls.Add(this.labelFilterCaption);
            this.Controls.Add(this.buttonClose2);
            this.Controls.Add(this.splitContainer1);
            this.Name = "ReportControl";
            this.Size = new System.Drawing.Size(800, 429);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox8)).EndInit();
            this.panelView.ResumeLayout(false);
            this.panelNativeView.ResumeLayout(false);
            this.panelNativeView.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TreeView treeViewReport;
        private System.Windows.Forms.TextBox textBoxReport;
        private System.Windows.Forms.Button buttonClose2;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label labelFilterCaption;
        private System.Windows.Forms.ComboBox comboBoxFilter;
        private System.Windows.Forms.PictureBox pictureBox8;
        private System.Windows.Forms.Label labelFilterHint;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.Panel panelView;
        private System.Windows.Forms.Panel panelNativeView;
        private System.Windows.Forms.CheckBox checkBoxNativeView;
    }
}
