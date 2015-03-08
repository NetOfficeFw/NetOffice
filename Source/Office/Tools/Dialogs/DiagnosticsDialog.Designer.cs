namespace NetOffice.OfficeApi.Tools.Dialogs
{
    partial class DiagnosticsDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DiagnosticsDialog));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.labelAssemblyInfo = new System.Windows.Forms.Label();
            this.pictureBoxHeader = new System.Windows.Forms.PictureBox();
            this.dataGridViewDiagnostics = new System.Windows.Forms.DataGridView();
            this.colType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.buttonClose = new System.Windows.Forms.Button();
            this.buttonClipboardCopy = new System.Windows.Forms.Button();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDiagnostics)).BeginInit();
            this.SuspendLayout();
            // 
            // panelHeader
            // 
            this.panelHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelHeader.BackColor = System.Drawing.Color.White;
            this.panelHeader.Controls.Add(this.labelAssemblyInfo);
            this.panelHeader.Controls.Add(this.pictureBoxHeader);
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(533, 58);
            this.panelHeader.TabIndex = 0;
            // 
            // labelAssemblyInfo
            // 
            this.labelAssemblyInfo.AutoSize = true;
            this.labelAssemblyInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelAssemblyInfo.ForeColor = System.Drawing.Color.Black;
            this.labelAssemblyInfo.Location = new System.Drawing.Point(73, 20);
            this.labelAssemblyInfo.Name = "labelAssemblyInfo";
            this.labelAssemblyInfo.Size = new System.Drawing.Size(165, 16);
            this.labelAssemblyInfo.TabIndex = 1;
            this.labelAssemblyInfo.Text = "Technical Environment";
            // 
            // pictureBoxHeader
            // 
            this.pictureBoxHeader.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxHeader.Image")));
            this.pictureBoxHeader.Location = new System.Drawing.Point(25, 13);
            this.pictureBoxHeader.Name = "pictureBoxHeader";
            this.pictureBoxHeader.Size = new System.Drawing.Size(34, 34);
            this.pictureBoxHeader.TabIndex = 0;
            this.pictureBoxHeader.TabStop = false;
            // 
            // dataGridViewDiagnostics
            // 
            this.dataGridViewDiagnostics.AllowUserToAddRows = false;
            this.dataGridViewDiagnostics.AllowUserToDeleteRows = false;
            this.dataGridViewDiagnostics.AllowUserToResizeRows = false;
            this.dataGridViewDiagnostics.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewDiagnostics.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.LightSteelBlue;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.Blue;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewDiagnostics.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewDiagnostics.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewDiagnostics.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colType,
            this.colValue});
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewDiagnostics.DefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridViewDiagnostics.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridViewDiagnostics.Location = new System.Drawing.Point(21, 82);
            this.dataGridViewDiagnostics.Name = "dataGridViewDiagnostics";
            this.dataGridViewDiagnostics.RowHeadersVisible = false;
            this.dataGridViewDiagnostics.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewDiagnostics.Size = new System.Drawing.Size(487, 209);
            this.dataGridViewDiagnostics.TabIndex = 1;
            // 
            // colType
            // 
            this.colType.DataPropertyName = "Type";
            this.colType.HeaderText = "Type";
            this.colType.Name = "colType";
            this.colType.Width = 150;
            // 
            // colValue
            // 
            this.colValue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colValue.DataPropertyName = "Value";
            this.colValue.HeaderText = "Value";
            this.colValue.Name = "colValue";
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonClose.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.ForeColor = System.Drawing.Color.Blue;
            this.buttonClose.Image = ((System.Drawing.Image)(resources.GetObject("buttonClose.Image")));
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.Location = new System.Drawing.Point(332, 310);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(176, 29);
            this.buttonClose.TabIndex = 2;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // buttonClipboardCopy
            // 
            this.buttonClipboardCopy.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonClipboardCopy.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonClipboardCopy.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClipboardCopy.ForeColor = System.Drawing.Color.Blue;
            this.buttonClipboardCopy.Image = ((System.Drawing.Image)(resources.GetObject("buttonClipboardCopy.Image")));
            this.buttonClipboardCopy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClipboardCopy.Location = new System.Drawing.Point(21, 310);
            this.buttonClipboardCopy.Name = "buttonClipboardCopy";
            this.buttonClipboardCopy.Size = new System.Drawing.Size(219, 29);
            this.buttonClipboardCopy.TabIndex = 3;
            this.buttonClipboardCopy.Text = "Copy to Clipboard";
            this.buttonClipboardCopy.UseVisualStyleBackColor = true;
            this.buttonClipboardCopy.Click += new System.EventHandler(this.buttonClipboardCopy_Click);
            // 
            // DiagnosticsDialog
            // 
            this.AcceptButton = this.buttonClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.CancelButton = this.buttonClose;
            this.ClientSize = new System.Drawing.Size(532, 353);
            this.Controls.Add(this.buttonClipboardCopy);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.dataGridViewDiagnostics);
            this.Controls.Add(this.panelHeader);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.MinimumSize = new System.Drawing.Size(540, 380);
            this.Name = "DiagnosticsDialog";
            this.Text = "Diagnostics";
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDiagnostics)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.PictureBox pictureBoxHeader;
        private System.Windows.Forms.Label labelAssemblyInfo;
        private System.Windows.Forms.DataGridView dataGridViewDiagnostics;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Button buttonClipboardCopy;
        private System.Windows.Forms.DataGridViewTextBoxColumn colType;
        private System.Windows.Forms.DataGridViewTextBoxColumn colValue;
    }
}