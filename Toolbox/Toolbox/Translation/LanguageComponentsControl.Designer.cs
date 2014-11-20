namespace NetOffice.DeveloperToolbox.Translation
{
    partial class LanguageComponentsControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LanguageComponentsControl));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelHint = new System.Windows.Forms.Label();
            this.textBoxRichString = new NetOffice.DeveloperToolbox.Controls.Text.RichTextEditor();
            this.textBoxWideString = new System.Windows.Forms.TextBox();
            this.textBoxString = new System.Windows.Forms.TextBox();
            this.imageStrip = new System.Windows.Forms.ImageList(this.components);
            this.treeGridView1 = new NetOffice.DeveloperToolbox.Controls.Tree.TreeGridView();
            this.ColumnName = new NetOffice.DeveloperToolbox.Controls.Tree.TreeGridColumn();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.treeGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.HotTrack = true;
            this.tabControl1.Location = new System.Drawing.Point(213, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(734, 568);
            this.tabControl1.TabIndex = 3;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(726, 539);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "View [ReadOnly]";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tabPage2.Controls.Add(this.textBoxRichString);
            this.tabPage2.Controls.Add(this.textBoxWideString);
            this.tabPage2.Controls.Add(this.textBoxString);
            this.tabPage2.Controls.Add(this.pictureBox1);
            this.tabPage2.Controls.Add(this.labelHint);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(726, 539);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "String Editor";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(17, 512);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(19, 19);
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // labelHint
            // 
            this.labelHint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelHint.AutoSize = true;
            this.labelHint.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelHint.ForeColor = System.Drawing.Color.DarkGoldenrod;
            this.labelHint.Location = new System.Drawing.Point(42, 508);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(510, 21);
            this.labelHint.TabIndex = 3;
            this.labelHint.Text = "To switch between elements use also Alt + Arrow(Up+Down) keys";
            // 
            // textBoxRichString
            // 
            this.textBoxRichString.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxRichString.Location = new System.Drawing.Point(31, 302);
            this.textBoxRichString.Name = "textBoxRichString";
            this.textBoxRichString.RichText = "{\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang1031{\\fonttbl{\\f0\\fnil\\fcharset0 Microsoft S" +
                "ans Serif;}}\r\n\\viewkind4\\uc1\\pard\\f0\\fs17\\par\r\n}\r\n";
            this.textBoxRichString.Size = new System.Drawing.Size(555, 114);
            this.textBoxRichString.TabIndex = 2;
            this.textBoxRichString.Visible = false;
            this.textBoxRichString.TextChanged += new System.EventHandler(this.textBoxRichString_TextChanged);
            // 
            // textBoxWideString
            // 
            this.textBoxWideString.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxWideString.Location = new System.Drawing.Point(31, 141);
            this.textBoxWideString.Multiline = true;
            this.textBoxWideString.Name = "textBoxWideString";
            this.textBoxWideString.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxWideString.Size = new System.Drawing.Size(555, 104);
            this.textBoxWideString.TabIndex = 1;
            this.textBoxWideString.Visible = false;
            this.textBoxWideString.TextChanged += new System.EventHandler(this.textBoxWideString_TextChanged);
            // 
            // textBoxString
            // 
            this.textBoxString.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxString.Location = new System.Drawing.Point(31, 87);
            this.textBoxString.Name = "textBoxString";
            this.textBoxString.Size = new System.Drawing.Size(555, 20);
            this.textBoxString.TabIndex = 0;
            this.textBoxString.Visible = false;
            this.textBoxString.TextChanged += new System.EventHandler(this.textBoxString_TextChanged);
            // 
            // imageStrip
            // 
            this.imageStrip.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageStrip.ImageStream")));
            this.imageStrip.TransparentColor = System.Drawing.Color.Transparent;
            this.imageStrip.Images.SetKeyName(0, "window_dialog.png");
            this.imageStrip.Images.SetKeyName(1, "text.png");
            this.imageStrip.Images.SetKeyName(2, "text_rich.png");
            this.imageStrip.Images.SetKeyName(3, "text_rich_colored.png");
            // 
            // treeGridView1
            // 
            this.treeGridView1.AllowUserToAddRows = false;
            this.treeGridView1.AllowUserToDeleteRows = false;
            this.treeGridView1.AllowUserToResizeColumns = false;
            this.treeGridView1.AllowUserToResizeRows = false;
            this.treeGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.treeGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.treeGridView1.BackgroundColor = System.Drawing.Color.LightSteelBlue;
            this.treeGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.treeGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnName});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.Orange;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Blue;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.treeGridView1.DefaultCellStyle = dataGridViewCellStyle1;
            this.treeGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.treeGridView1.GridColor = System.Drawing.Color.Blue;
            this.treeGridView1.ImageList = this.imageStrip;
            this.treeGridView1.Location = new System.Drawing.Point(1, 0);
            this.treeGridView1.MultiSelect = false;
            this.treeGridView1.Name = "treeGridView1";
            this.treeGridView1.RowHeadersVisible = false;
            this.treeGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.treeGridView1.ShowCellErrors = false;
            this.treeGridView1.ShowCellToolTips = false;
            this.treeGridView1.ShowEditingIcon = false;
            this.treeGridView1.ShowLines = false;
            this.treeGridView1.ShowRowErrors = false;
            this.treeGridView1.Size = new System.Drawing.Size(210, 564);
            this.treeGridView1.TabIndex = 4;
            this.treeGridView1.SelectionChanged += new System.EventHandler(this.treeGridView1_SelectionChanged);
            this.treeGridView1.DoubleClick += new System.EventHandler(this.treeGridView1_DoubleClick);
            // 
            // ColumnName
            // 
            this.ColumnName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ColumnName.DefaultNodeImage = null;
            this.ColumnName.HeaderText = "Components";
            this.ColumnName.Name = "ColumnName";
            this.ColumnName.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColumnName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // LanguageComponentsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.treeGridView1);
            this.Controls.Add(this.tabControl1);
            this.Name = "LanguageComponentsControl";
            this.Size = new System.Drawing.Size(948, 568);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.treeGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private Controls.Tree.TreeGridView treeGridView1;
        private System.Windows.Forms.ImageList imageStrip;
        private Controls.Tree.TreeGridColumn ColumnName;
        private System.Windows.Forms.TextBox textBoxString;
        private System.Windows.Forms.TextBox textBoxWideString;
        private Controls.Text.RichTextEditor textBoxRichString;
        private System.Windows.Forms.Label labelHint;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}
