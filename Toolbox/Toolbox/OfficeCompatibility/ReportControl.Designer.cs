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
            this.label1 = new System.Windows.Forms.Label();
            this.buttonSaveReport = new System.Windows.Forms.Button();
            this.panelNativeView = new System.Windows.Forms.Panel();
            this.textBoxReport = new System.Windows.Forms.TextBox();
            this.panelView = new System.Windows.Forms.Panel();
            this.listView1 = new System.Windows.Forms.ListView();
            this.checkBoxNativeView = new System.Windows.Forms.CheckBox();
            this.buttonClose2 = new System.Windows.Forms.Button();
            this.labelFilterCaption = new System.Windows.Forms.Label();
            this.comboBoxFilter = new System.Windows.Forms.ComboBox();
            this.labelFilterHint = new System.Windows.Forms.Label();
            this.pictureBoxField = new System.Windows.Forms.PictureBox();
            this.labelFieldHint = new System.Windows.Forms.Label();
            this.labelPropertyHint = new System.Windows.Forms.Label();
            this.pictureBoxProperty = new System.Windows.Forms.PictureBox();
            this.labelMethodHint = new System.Windows.Forms.Label();
            this.pictureBoxMethod = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.panelNativeView.SuspendLayout();
            this.panelView.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxField)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxProperty)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxMethod)).BeginInit();
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
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            this.splitContainer1.Panel1.Controls.Add(this.buttonSaveReport);
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
            this.treeViewReport.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.treeViewReport.HideSelection = false;
            this.treeViewReport.ImageIndex = 0;
            this.treeViewReport.ImageList = this.imageList1;
            this.treeViewReport.Location = new System.Drawing.Point(0, 15);
            this.treeViewReport.Name = "treeViewReport";
            this.treeViewReport.SelectedImageIndex = 0;
            this.treeViewReport.Size = new System.Drawing.Size(265, 340);
            this.treeViewReport.TabIndex = 0;
            this.treeViewReport.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeViewReport_AfterSelect);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Magenta;
            this.imageList1.Images.SetKeyName(0, "VSObject_Assembly.bmp");
            this.imageList1.Images.SetKeyName(1, "VSObject_Class.bmp");
            this.imageList1.Images.SetKeyName(2, "VSObject_Class_Private.bmp");
            this.imageList1.Images.SetKeyName(3, "VSObject_Field.bmp");
            this.imageList1.Images.SetKeyName(4, "VSObject_Field_Private.bmp");
            this.imageList1.Images.SetKeyName(5, "VSObject_Method.bmp");
            this.imageList1.Images.SetKeyName(6, "VSObject_Method_Private.bmp");
            this.imageList1.Images.SetKeyName(7, "VSObject_Properties.bmp");
            this.imageList1.Images.SetKeyName(8, "VSObject_Properties_Private.bmp");
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.BackColor = System.Drawing.Color.Khaki;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(265, 21);
            this.label1.TabIndex = 75;
            this.label1.Text = "Assembly Explorer";
            // 
            // buttonSaveReport
            // 
            this.buttonSaveReport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSaveReport.Location = new System.Drawing.Point(2, 358);
            this.buttonSaveReport.Name = "buttonSaveReport";
            this.buttonSaveReport.Size = new System.Drawing.Size(262, 24);
            this.buttonSaveReport.TabIndex = 74;
            this.buttonSaveReport.Text = "Bericht in Datei speichern";
            this.buttonSaveReport.UseVisualStyleBackColor = true;
            this.buttonSaveReport.Click += new System.EventHandler(this.buttonSaveReport_Click);
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
            // panelView
            // 
            this.panelView.Controls.Add(this.listView1);
            this.panelView.Location = new System.Drawing.Point(28, 12);
            this.panelView.Name = "panelView";
            this.panelView.Size = new System.Drawing.Size(497, 168);
            this.panelView.TabIndex = 0;
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
            this.listView1.Resize += new System.EventHandler(this.listView1_Resize);
            // 
            // checkBoxNativeView
            // 
            this.checkBoxNativeView.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBoxNativeView.AutoSize = true;
            this.checkBoxNativeView.Location = new System.Drawing.Point(652, 0);
            this.checkBoxNativeView.Name = "checkBoxNativeView";
            this.checkBoxNativeView.Size = new System.Drawing.Size(95, 17);
            this.checkBoxNativeView.TabIndex = 73;
            this.checkBoxNativeView.Text = "Native Ansicht";
            this.checkBoxNativeView.UseVisualStyleBackColor = true;
            this.checkBoxNativeView.Visible = false;
            this.checkBoxNativeView.CheckedChanged += new System.EventHandler(this.checkBoxNativeView_CheckedChanged);
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
            this.labelFilterCaption.Location = new System.Drawing.Point(270, 23);
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
            this.comboBoxFilter.Location = new System.Drawing.Point(348, 17);
            this.comboBoxFilter.Name = "comboBoxFilter";
            this.comboBoxFilter.Size = new System.Drawing.Size(123, 21);
            this.comboBoxFilter.TabIndex = 32;
            this.comboBoxFilter.SelectedIndexChanged += new System.EventHandler(this.comboBoxFilter_SelectedIndexChanged);
            // 
            // labelFilterHint
            // 
            this.labelFilterHint.BackColor = System.Drawing.Color.Khaki;
            this.labelFilterHint.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelFilterHint.Location = new System.Drawing.Point(474, 12);
            this.labelFilterHint.Name = "labelFilterHint";
            this.labelFilterHint.Size = new System.Drawing.Size(276, 26);
            this.labelFilterHint.TabIndex = 71;
            this.labelFilterHint.Text = "Nutzen Sie den Filter um nur Funktionalitäten anzuzeigen die eine bestimmte Versi" +
                "on nicht unterstützen.";
            // 
            // pictureBoxField
            // 
            this.pictureBoxField.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxField.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxField.Image")));
            this.pictureBoxField.Location = new System.Drawing.Point(5, 22);
            this.pictureBoxField.Name = "pictureBoxField";
            this.pictureBoxField.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxField.TabIndex = 74;
            this.pictureBoxField.TabStop = false;
            // 
            // labelFieldHint
            // 
            this.labelFieldHint.AutoSize = true;
            this.labelFieldHint.Location = new System.Drawing.Point(27, 23);
            this.labelFieldHint.Name = "labelFieldHint";
            this.labelFieldHint.Size = new System.Drawing.Size(38, 13);
            this.labelFieldHint.TabIndex = 75;
            this.labelFieldHint.Text = "= Field";
            // 
            // labelPropertyHint
            // 
            this.labelPropertyHint.AutoSize = true;
            this.labelPropertyHint.Location = new System.Drawing.Point(100, 23);
            this.labelPropertyHint.Name = "labelPropertyHint";
            this.labelPropertyHint.Size = new System.Drawing.Size(55, 13);
            this.labelPropertyHint.TabIndex = 77;
            this.labelPropertyHint.Text = "= Property";
            // 
            // pictureBoxProperty
            // 
            this.pictureBoxProperty.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxProperty.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxProperty.Image")));
            this.pictureBoxProperty.Location = new System.Drawing.Point(78, 22);
            this.pictureBoxProperty.Name = "pictureBoxProperty";
            this.pictureBoxProperty.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxProperty.TabIndex = 76;
            this.pictureBoxProperty.TabStop = false;
            // 
            // labelMethodHint
            // 
            this.labelMethodHint.AutoSize = true;
            this.labelMethodHint.Location = new System.Drawing.Point(185, 23);
            this.labelMethodHint.Name = "labelMethodHint";
            this.labelMethodHint.Size = new System.Drawing.Size(52, 13);
            this.labelMethodHint.TabIndex = 79;
            this.labelMethodHint.Text = "= Method";
            // 
            // pictureBoxMethod
            // 
            this.pictureBoxMethod.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxMethod.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxMethod.Image")));
            this.pictureBoxMethod.Location = new System.Drawing.Point(163, 22);
            this.pictureBoxMethod.Name = "pictureBoxMethod";
            this.pictureBoxMethod.Size = new System.Drawing.Size(16, 16);
            this.pictureBoxMethod.TabIndex = 78;
            this.pictureBoxMethod.TabStop = false;
            // 
            // ReportControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelFilterCaption);
            this.Controls.Add(this.labelMethodHint);
            this.Controls.Add(this.checkBoxNativeView);
            this.Controls.Add(this.comboBoxFilter);
            this.Controls.Add(this.pictureBoxMethod);
            this.Controls.Add(this.labelFilterHint);
            this.Controls.Add(this.labelPropertyHint);
            this.Controls.Add(this.pictureBoxProperty);
            this.Controls.Add(this.labelFieldHint);
            this.Controls.Add(this.pictureBoxField);
            this.Controls.Add(this.buttonClose2);
            this.Controls.Add(this.splitContainer1);
            this.Name = "ReportControl";
            this.Size = new System.Drawing.Size(800, 429);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.panelNativeView.ResumeLayout(false);
            this.panelNativeView.PerformLayout();
            this.panelView.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxField)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxProperty)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxMethod)).EndInit();
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
        private System.Windows.Forms.Label labelFilterHint;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.Panel panelView;
        private System.Windows.Forms.Panel panelNativeView;
        private System.Windows.Forms.CheckBox checkBoxNativeView;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonSaveReport;
        private System.Windows.Forms.PictureBox pictureBoxField;
        private System.Windows.Forms.Label labelFieldHint;
        private System.Windows.Forms.Label labelPropertyHint;
        private System.Windows.Forms.PictureBox pictureBoxProperty;
        private System.Windows.Forms.Label labelMethodHint;
        private System.Windows.Forms.PictureBox pictureBoxMethod;
    }
}
