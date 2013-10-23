namespace NOTools.CSharpTextEditor
{
    partial class ReferencePanel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReferencePanel));
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItemAdd = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemRemove = new System.Windows.Forms.ToolStripMenuItem();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.buttonOpenHide = new System.Windows.Forms.Button();
            this.buttonHide = new System.Windows.Forms.Button();
            this.buttonOpen = new System.Windows.Forms.Button();
            this.labelHeader = new System.Windows.Forms.Label();
            this.labelLine = new System.Windows.Forms.Label();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.treeView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.treeView1.ContextMenuStrip = this.contextMenuStrip1;
            this.treeView1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeView1.FullRowSelect = true;
            this.treeView1.ImageIndex = 0;
            this.treeView1.ImageList = this.imageList1;
            this.treeView1.Indent = 19;
            this.treeView1.ItemHeight = 20;
            this.treeView1.Location = new System.Drawing.Point(0, 17);
            this.treeView1.Name = "treeView1";
            this.treeView1.SelectedImageIndex = 0;
            this.treeView1.ShowLines = false;
            this.treeView1.ShowNodeToolTips = true;
            this.treeView1.ShowPlusMinus = false;
            this.treeView1.Size = new System.Drawing.Size(372, 340);
            this.treeView1.TabIndex = 0;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemAdd,
            this.toolStripMenuItemRemove});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(125, 48);
            this.contextMenuStrip1.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuStrip1_Opening);
            // 
            // toolStripMenuItemAdd
            // 
            this.toolStripMenuItemAdd.Image = ((System.Drawing.Image)(resources.GetObject("toolStripMenuItemAdd.Image")));
            this.toolStripMenuItemAdd.Name = "toolStripMenuItemAdd";
            this.toolStripMenuItemAdd.Size = new System.Drawing.Size(124, 22);
            this.toolStripMenuItemAdd.Text = "Add";
            this.toolStripMenuItemAdd.Click += new System.EventHandler(this.toolStripMenuItemAdd_Click);
            // 
            // toolStripMenuItemRemove
            // 
            this.toolStripMenuItemRemove.Image = ((System.Drawing.Image)(resources.GetObject("toolStripMenuItemRemove.Image")));
            this.toolStripMenuItemRemove.Name = "toolStripMenuItemRemove";
            this.toolStripMenuItemRemove.Size = new System.Drawing.Size(124, 22);
            this.toolStripMenuItemRemove.Text = "Remove";
            this.toolStripMenuItemRemove.Click += new System.EventHandler(this.toolStripMenuItemRemove_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "element.png");
            // 
            // buttonOpenHide
            // 
            this.buttonOpenHide.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOpenHide.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonOpenHide.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOpenHide.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpenHide.Image")));
            this.buttonOpenHide.Location = new System.Drawing.Point(356, 0);
            this.buttonOpenHide.Name = "buttonOpenHide";
            this.buttonOpenHide.Size = new System.Drawing.Size(17, 17);
            this.buttonOpenHide.TabIndex = 1;
            this.buttonOpenHide.UseVisualStyleBackColor = true;
            this.buttonOpenHide.Click += new System.EventHandler(this.buttonOpenHide_Click);
            // 
            // buttonHide
            // 
            this.buttonHide.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonHide.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonHide.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonHide.Image = ((System.Drawing.Image)(resources.GetObject("buttonHide.Image")));
            this.buttonHide.Location = new System.Drawing.Point(167, 0);
            this.buttonHide.Name = "buttonHide";
            this.buttonHide.Size = new System.Drawing.Size(17, 17);
            this.buttonHide.TabIndex = 5;
            this.buttonHide.UseVisualStyleBackColor = false;
            this.buttonHide.Visible = false;
            // 
            // buttonOpen
            // 
            this.buttonOpen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOpen.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOpen.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpen.Image")));
            this.buttonOpen.Location = new System.Drawing.Point(190, 0);
            this.buttonOpen.Name = "buttonOpen";
            this.buttonOpen.Size = new System.Drawing.Size(17, 17);
            this.buttonOpen.TabIndex = 4;
            this.buttonOpen.UseVisualStyleBackColor = false;
            this.buttonOpen.Visible = false;
            // 
            // labelHeader
            // 
            this.labelHeader.AutoSize = true;
            this.labelHeader.Location = new System.Drawing.Point(8, 2);
            this.labelHeader.Name = "labelHeader";
            this.labelHeader.Size = new System.Drawing.Size(62, 13);
            this.labelHeader.TabIndex = 6;
            this.labelHeader.Text = "References";
            // 
            // labelLine
            // 
            this.labelLine.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelLine.BackColor = System.Drawing.Color.Black;
            this.labelLine.ForeColor = System.Drawing.Color.Black;
            this.labelLine.Location = new System.Drawing.Point(0, 16);
            this.labelLine.Name = "labelLine";
            this.labelLine.Size = new System.Drawing.Size(356, 1);
            this.labelLine.TabIndex = 7;
            this.labelLine.Text = "label2";
            // 
            // ReferencePanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.labelLine);
            this.Controls.Add(this.labelHeader);
            this.Controls.Add(this.buttonHide);
            this.Controls.Add(this.buttonOpen);
            this.Controls.Add(this.buttonOpenHide);
            this.Controls.Add(this.treeView1);
            this.Name = "ReferencePanel";
            this.Size = new System.Drawing.Size(373, 357);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button buttonOpenHide;
        private System.Windows.Forms.Button buttonHide;
        private System.Windows.Forms.Button buttonOpen;
        private System.Windows.Forms.Label labelLine;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemAdd;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemRemove;
        internal System.Windows.Forms.Label labelHeader;
    }
}
