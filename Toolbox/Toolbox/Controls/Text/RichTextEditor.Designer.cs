namespace NetOffice.DeveloperToolbox.Controls.Text
{
    partial class RichTextEditor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RichTextEditor));
            this.Toolbox = new System.Windows.Forms.ToolStrip();
            this.toolStripComboBox1 = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripComboBox2 = new System.Windows.Forms.ToolStripComboBox();
            this.sepTBFormatting1 = new System.Windows.Forms.ToolStripSeparator();
            this.buttonBold = new System.Windows.Forms.ToolStripButton();
            this.buttonItalic = new System.Windows.Forms.ToolStripButton();
            this.buttonUnderline = new System.Windows.Forms.ToolStripButton();
            this.buttonStrikeout = new System.Windows.Forms.ToolStripButton();
            this.sepTBFormatting2 = new System.Windows.Forms.ToolStripSeparator();
            this.buttonForeColor = new System.Windows.Forms.ToolStripButton();
            this.buttonBackColor = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.buttonImport = new System.Windows.Forms.ToolStripButton();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.labelHint = new System.Windows.Forms.Label();
            this.Toolbox.SuspendLayout();
            this.SuspendLayout();
            // 
            // Toolbox
            // 
            this.Toolbox.BackColor = System.Drawing.Color.Transparent;
            this.Toolbox.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripComboBox1,
            this.toolStripComboBox2,
            this.sepTBFormatting1,
            this.buttonBold,
            this.buttonItalic,
            this.buttonUnderline,
            this.buttonStrikeout,
            this.sepTBFormatting2,
            this.buttonForeColor,
            this.buttonBackColor,
            this.toolStripSeparator1,
            this.buttonImport});
            this.Toolbox.Location = new System.Drawing.Point(0, 0);
            this.Toolbox.Name = "Toolbox";
            this.Toolbox.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.Toolbox.Size = new System.Drawing.Size(555, 25);
            this.Toolbox.TabIndex = 5;
            this.Toolbox.Text = "toolStrip1";
            // 
            // toolStripComboBox1
            // 
            this.toolStripComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.toolStripComboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.toolStripComboBox1.DropDownHeight = 300;
            this.toolStripComboBox1.IntegralHeight = false;
            this.toolStripComboBox1.Name = "toolStripComboBox1";
            this.toolStripComboBox1.Size = new System.Drawing.Size(170, 25);
            this.toolStripComboBox1.SelectedIndexChanged += new System.EventHandler(this.toolStripComboBox1_SelectedIndexChanged);
            // 
            // toolStripComboBox2
            // 
            this.toolStripComboBox2.AutoSize = false;
            this.toolStripComboBox2.DropDownHeight = 200;
            this.toolStripComboBox2.IntegralHeight = false;
            this.toolStripComboBox2.Items.AddRange(new object[] {
            "8",
            "9",
            "10",
            "11",
            "12",
            "14"});
            this.toolStripComboBox2.MaxDropDownItems = 20;
            this.toolStripComboBox2.Name = "toolStripComboBox2";
            this.toolStripComboBox2.Size = new System.Drawing.Size(50, 21);
            this.toolStripComboBox2.Text = "8";
            this.toolStripComboBox2.SelectedIndexChanged += new System.EventHandler(this.toolStripComboBox2_SelectedIndexChanged);
            // 
            // sepTBFormatting1
            // 
            this.sepTBFormatting1.Name = "sepTBFormatting1";
            this.sepTBFormatting1.Size = new System.Drawing.Size(6, 25);
            // 
            // buttonBold
            // 
            this.buttonBold.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonBold.Image = ((System.Drawing.Image)(resources.GetObject("buttonBold.Image")));
            this.buttonBold.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonBold.Name = "buttonBold";
            this.buttonBold.Size = new System.Drawing.Size(23, 22);
            this.buttonBold.Text = "Bold";
            this.buttonBold.Click += new System.EventHandler(this.buttonBold_Click);
            // 
            // buttonItalic
            // 
            this.buttonItalic.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonItalic.Image = ((System.Drawing.Image)(resources.GetObject("buttonItalic.Image")));
            this.buttonItalic.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonItalic.Name = "buttonItalic";
            this.buttonItalic.Size = new System.Drawing.Size(23, 22);
            this.buttonItalic.Text = "Italic";
            this.buttonItalic.Click += new System.EventHandler(this.buttonItalic_Click);
            // 
            // buttonUnderline
            // 
            this.buttonUnderline.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonUnderline.Image = ((System.Drawing.Image)(resources.GetObject("buttonUnderline.Image")));
            this.buttonUnderline.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonUnderline.Name = "buttonUnderline";
            this.buttonUnderline.Size = new System.Drawing.Size(23, 22);
            this.buttonUnderline.Text = "Underline";
            this.buttonUnderline.Click += new System.EventHandler(this.buttonUnderline_Click);
            // 
            // buttonStrikeout
            // 
            this.buttonStrikeout.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonStrikeout.Image = ((System.Drawing.Image)(resources.GetObject("buttonStrikeout.Image")));
            this.buttonStrikeout.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonStrikeout.Name = "buttonStrikeout";
            this.buttonStrikeout.Size = new System.Drawing.Size(23, 22);
            this.buttonStrikeout.Text = "Strikeout";
            this.buttonStrikeout.Click += new System.EventHandler(this.buttonStrikeout_Click);
            // 
            // sepTBFormatting2
            // 
            this.sepTBFormatting2.Name = "sepTBFormatting2";
            this.sepTBFormatting2.Size = new System.Drawing.Size(6, 25);
            // 
            // buttonForeColor
            // 
            this.buttonForeColor.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.buttonForeColor.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonForeColor.Image = ((System.Drawing.Image)(resources.GetObject("buttonForeColor.Image")));
            this.buttonForeColor.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonForeColor.Name = "buttonForeColor";
            this.buttonForeColor.Size = new System.Drawing.Size(23, 22);
            this.buttonForeColor.Text = "A";
            this.buttonForeColor.ToolTipText = "Fore Color";
            this.buttonForeColor.Click += new System.EventHandler(this.buttonForeColor_Click);
            // 
            // buttonBackColor
            // 
            this.buttonBackColor.BackColor = System.Drawing.Color.White;
            this.buttonBackColor.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.buttonBackColor.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonBackColor.Image = ((System.Drawing.Image)(resources.GetObject("buttonBackColor.Image")));
            this.buttonBackColor.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonBackColor.Name = "buttonBackColor";
            this.buttonBackColor.Size = new System.Drawing.Size(23, 22);
            this.buttonBackColor.Text = "A";
            this.buttonBackColor.ToolTipText = "Back Color";
            this.buttonBackColor.Click += new System.EventHandler(this.buttonBackColor_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // buttonImport
            // 
            this.buttonImport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonImport.Image = ((System.Drawing.Image)(resources.GetObject("buttonImport.Image")));
            this.buttonImport.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonImport.Name = "buttonImport";
            this.buttonImport.Size = new System.Drawing.Size(23, 22);
            this.buttonImport.Text = "Load Rich Text File";
            this.buttonImport.Click += new System.EventHandler(this.buttonImport_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox1.Location = new System.Drawing.Point(0, 26);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(553, 221);
            this.richTextBox1.TabIndex = 6;
            this.richTextBox1.Text = "";
            this.richTextBox1.SelectionChanged += new System.EventHandler(this.richTextBox1_SelectionChanged);
            this.richTextBox1.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
            // 
            // labelHint
            // 
            this.labelHint.AutoSize = true;
            this.labelHint.ForeColor = System.Drawing.Color.Red;
            this.labelHint.Location = new System.Drawing.Point(419, 6);
            this.labelHint.Name = "labelHint";
            this.labelHint.Size = new System.Drawing.Size(309, 13);
            this.labelHint.TabIndex = 7;
            this.labelHint.Text = "Editor is not complete. Use import option to use all RTF features.";
            // 
            // RichTextEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelHint);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.Toolbox);
            this.Name = "RichTextEditor";
            this.Size = new System.Drawing.Size(555, 249);
            this.Toolbox.ResumeLayout(false);
            this.Toolbox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip Toolbox;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox1;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox2;
        private System.Windows.Forms.ToolStripSeparator sepTBFormatting1;
        private System.Windows.Forms.ToolStripButton buttonBold;
        private System.Windows.Forms.ToolStripButton buttonItalic;
        private System.Windows.Forms.ToolStripButton buttonUnderline;
        private System.Windows.Forms.ToolStripButton buttonStrikeout;
        private System.Windows.Forms.ToolStripSeparator sepTBFormatting2;
        private System.Windows.Forms.ToolStripButton buttonForeColor;
        private System.Windows.Forms.ToolStripButton buttonBackColor;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton buttonImport;
        private System.Windows.Forms.Label labelHint;
    }
}
