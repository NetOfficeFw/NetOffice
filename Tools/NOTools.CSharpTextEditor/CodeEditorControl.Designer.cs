namespace NOTools.CSharpTextEditor
{
    partial class CodeEditorControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CodeEditorControl));
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.wpfControl1 = new NOTools.CSharpTextEditor.WPFControl();
            this.referencePanel1 = new NOTools.CSharpTextEditor.ReferencePanel();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.labelInfo = new System.Windows.Forms.Label();
            this.buttonHide = new System.Windows.Forms.Button();
            this.buttonOpen = new System.Windows.Forms.Button();
            this.labelErrors = new System.Windows.Forms.Label();
            this.buttonErrorPanelOpenHide = new System.Windows.Forms.Button();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.errorPanel1 = new NOTools.CSharpTextEditor.ErrorPanel();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).BeginInit();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer3.Panel2.SuspendLayout();
            this.splitContainer3.SuspendLayout();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.BackColor = System.Drawing.Color.Black;
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.BackColor = System.Drawing.Color.Black;
            this.splitContainer1.Panel1.Controls.Add(this.splitContainer3);
            this.splitContainer1.Panel1.Controls.Add(this.panelHeader);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(400, 400);
            this.splitContainer1.SplitterDistance = 357;
            this.splitContainer1.SplitterWidth = 1;
            this.splitContainer1.TabIndex = 1;
            // 
            // splitContainer3
            // 
            this.splitContainer3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer3.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.splitContainer3.Location = new System.Drawing.Point(1, 1);
            this.splitContainer3.Name = "splitContainer3";
            // 
            // splitContainer3.Panel1
            // 
            this.splitContainer3.Panel1.Controls.Add(this.elementHost1);
            // 
            // splitContainer3.Panel2
            // 
            this.splitContainer3.Panel2.Controls.Add(this.referencePanel1);
            this.splitContainer3.Panel2MinSize = 10;
            this.splitContainer3.Size = new System.Drawing.Size(398, 336);
            this.splitContainer3.SplitterDistance = 235;
            this.splitContainer3.TabIndex = 4;
            // 
            // elementHost1
            // 
            this.elementHost1.BackColor = System.Drawing.Color.White;
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(235, 336);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.wpfControl1;
            // 
            // referencePanel1
            // 
            this.referencePanel1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.referencePanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.referencePanel1.Location = new System.Drawing.Point(0, 0);
            this.referencePanel1.Name = "referencePanel1";
            this.referencePanel1.Size = new System.Drawing.Size(159, 336);
            this.referencePanel1.TabIndex = 0;
            this.referencePanel1.OpenHideClick += new System.EventHandler(this.referencePanel1_OpenHideClick);
            // 
            // panelHeader
            // 
            this.panelHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelHeader.BackColor = System.Drawing.Color.LightSteelBlue;
            this.panelHeader.Controls.Add(this.labelInfo);
            this.panelHeader.Controls.Add(this.buttonHide);
            this.panelHeader.Controls.Add(this.buttonOpen);
            this.panelHeader.Controls.Add(this.labelErrors);
            this.panelHeader.Controls.Add(this.buttonErrorPanelOpenHide);
            this.panelHeader.Location = new System.Drawing.Point(1, 338);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(398, 17);
            this.panelHeader.TabIndex = 3;
            // 
            // labelInfo
            // 
            this.labelInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelInfo.AutoSize = true;
            this.labelInfo.Location = new System.Drawing.Point(206, 2);
            this.labelInfo.Name = "labelInfo";
            this.labelInfo.Size = new System.Drawing.Size(0, 13);
            this.labelInfo.TabIndex = 4;
            this.labelInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // buttonHide
            // 
            this.buttonHide.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonHide.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonHide.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonHide.Image = ((System.Drawing.Image)(resources.GetObject("buttonHide.Image")));
            this.buttonHide.Location = new System.Drawing.Point(160, 1);
            this.buttonHide.Name = "buttonHide";
            this.buttonHide.Size = new System.Drawing.Size(17, 17);
            this.buttonHide.TabIndex = 3;
            this.buttonHide.UseVisualStyleBackColor = false;
            this.buttonHide.Visible = false;
            // 
            // buttonOpen
            // 
            this.buttonOpen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOpen.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOpen.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpen.Image")));
            this.buttonOpen.Location = new System.Drawing.Point(183, 2);
            this.buttonOpen.Name = "buttonOpen";
            this.buttonOpen.Size = new System.Drawing.Size(17, 17);
            this.buttonOpen.TabIndex = 2;
            this.buttonOpen.UseVisualStyleBackColor = false;
            this.buttonOpen.Visible = false;
            // 
            // labelErrors
            // 
            this.labelErrors.AutoSize = true;
            this.labelErrors.Location = new System.Drawing.Point(8, 2);
            this.labelErrors.Name = "labelErrors";
            this.labelErrors.Size = new System.Drawing.Size(34, 13);
            this.labelErrors.TabIndex = 1;
            this.labelErrors.Text = "Errors";
            // 
            // buttonErrorPanelOpenHide
            // 
            this.buttonErrorPanelOpenHide.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonErrorPanelOpenHide.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonErrorPanelOpenHide.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonErrorPanelOpenHide.Image = ((System.Drawing.Image)(resources.GetObject("buttonErrorPanelOpenHide.Image")));
            this.buttonErrorPanelOpenHide.Location = new System.Drawing.Point(381, 0);
            this.buttonErrorPanelOpenHide.Name = "buttonErrorPanelOpenHide";
            this.buttonErrorPanelOpenHide.Size = new System.Drawing.Size(17, 17);
            this.buttonErrorPanelOpenHide.TabIndex = 0;
            this.buttonErrorPanelOpenHide.UseVisualStyleBackColor = true;
            this.buttonErrorPanelOpenHide.Visible = false;
            this.buttonErrorPanelOpenHide.Click += new System.EventHandler(this.buttonErrorPanelOpenHide_Click);
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.errorPanel1);
            this.splitContainer2.Panel2Collapsed = true;
            this.splitContainer2.Size = new System.Drawing.Size(400, 42);
            this.splitContainer2.SplitterDistance = 263;
            this.splitContainer2.TabIndex = 1;
            // 
            // errorPanel1
            // 
            this.errorPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.errorPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.errorPanel1.ErrorColumnHeader = "Error";
            this.errorPanel1.LineColumnHeader = "Line";
            this.errorPanel1.Location = new System.Drawing.Point(0, 0);
            this.errorPanel1.Name = "errorPanel1";
            this.errorPanel1.Size = new System.Drawing.Size(400, 41);
            this.errorPanel1.TabIndex = 0;
            this.errorPanel1.ErrorDoubleClick += new NOTools.CSharpTextEditor.ErrorDoubleClickEventHandler(this.errorPanel1_ErrorDoubleClick);
            // 
            // CodeEditorControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.Controls.Add(this.splitContainer1);
            this.MinimumSize = new System.Drawing.Size(200, 200);
            this.Name = "CodeEditorControl";
            this.Size = new System.Drawing.Size(400, 400);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).EndInit();
            this.splitContainer3.ResumeLayout(false);
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            this.splitContainer2.Panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer2;
        internal ErrorPanel errorPanel1;
        internal System.Windows.Forms.Button buttonErrorPanelOpenHide;
        internal System.Windows.Forms.Label labelErrors;
        internal System.Windows.Forms.Button buttonOpen;
        internal System.Windows.Forms.Button buttonHide;
        internal System.Windows.Forms.SplitContainer splitContainer1;
        internal System.Windows.Forms.Label labelInfo;
        internal System.Windows.Forms.Panel panelHeader;
        internal System.Windows.Forms.SplitContainer splitContainer3;
        internal System.Windows.Forms.Integration.ElementHost elementHost1;
        internal WPFControl wpfControl1;
        internal ReferencePanel referencePanel1;
    }
}
