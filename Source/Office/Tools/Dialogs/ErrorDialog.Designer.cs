namespace NetOffice.OfficeApi.Tools.Dialogs
{
    partial class ErrorDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorDialog));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.labelErrorHeader = new System.Windows.Forms.Label();
            this.pictureBoxHeader = new System.Windows.Forms.PictureBox();
            this.labelErrorMessage = new System.Windows.Forms.Label();
            this.buttonClipboardCopy = new System.Windows.Forms.Button();
            this.dataGridViewErrors = new System.Windows.Forms.DataGridView();
            this.buttonShowDetails = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.colMessage = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSource = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewErrors)).BeginInit();
            this.SuspendLayout();
            // 
            // panelHeader
            // 
            this.panelHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelHeader.BackColor = System.Drawing.Color.White;
            this.panelHeader.Controls.Add(this.labelErrorHeader);
            this.panelHeader.Controls.Add(this.pictureBoxHeader);
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(536, 58);
            this.panelHeader.TabIndex = 1;
            // 
            // labelErrorHeader
            // 
            this.labelErrorHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorHeader.ForeColor = System.Drawing.Color.Black;
            this.labelErrorHeader.Location = new System.Drawing.Point(67, 22);
            this.labelErrorHeader.Name = "labelErrorHeader";
            this.labelErrorHeader.Size = new System.Drawing.Size(462, 33);
            this.labelErrorHeader.TabIndex = 1;
            this.labelErrorHeader.Text = "%ErrorNotification";
            // 
            // pictureBoxHeader
            // 
            this.pictureBoxHeader.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxHeader.Image")));
            this.pictureBoxHeader.Location = new System.Drawing.Point(24, 13);
            this.pictureBoxHeader.Name = "pictureBoxHeader";
            this.pictureBoxHeader.Size = new System.Drawing.Size(34, 34);
            this.pictureBoxHeader.TabIndex = 0;
            this.pictureBoxHeader.TabStop = false;
            // 
            // labelErrorMessage
            // 
            this.labelErrorMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorMessage.ForeColor = System.Drawing.Color.Black;
            this.labelErrorMessage.Location = new System.Drawing.Point(28, 80);
            this.labelErrorMessage.Name = "labelErrorMessage";
            this.labelErrorMessage.Size = new System.Drawing.Size(479, 43);
            this.labelErrorMessage.TabIndex = 2;
            this.labelErrorMessage.Text = "%ErrorMessage";
            // 
            // buttonClipboardCopy
            // 
            this.buttonClipboardCopy.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClipboardCopy.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonClipboardCopy.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClipboardCopy.ForeColor = System.Drawing.Color.Blue;
            this.buttonClipboardCopy.Image = ((System.Drawing.Image)(resources.GetObject("buttonClipboardCopy.Image")));
            this.buttonClipboardCopy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClipboardCopy.Location = new System.Drawing.Point(24, 428);
            this.buttonClipboardCopy.Name = "buttonClipboardCopy";
            this.buttonClipboardCopy.Size = new System.Drawing.Size(486, 29);
            this.buttonClipboardCopy.TabIndex = 6;
            this.buttonClipboardCopy.Text = "Copy to Clipboard";
            this.buttonClipboardCopy.UseVisualStyleBackColor = true;
            this.buttonClipboardCopy.Visible = false;
            this.buttonClipboardCopy.Click += new System.EventHandler(this.buttonClipboardCopy_Click);
            // 
            // dataGridViewErrors
            // 
            this.dataGridViewErrors.AllowUserToAddRows = false;
            this.dataGridViewErrors.AllowUserToDeleteRows = false;
            this.dataGridViewErrors.AllowUserToOrderColumns = true;
            this.dataGridViewErrors.AllowUserToResizeRows = false;
            this.dataGridViewErrors.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewErrors.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.LightSteelBlue;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Blue;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewErrors.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewErrors.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewErrors.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colMessage,
            this.colType,
            this.colSource});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewErrors.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewErrors.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridViewErrors.Location = new System.Drawing.Point(23, 190);
            this.dataGridViewErrors.MultiSelect = false;
            this.dataGridViewErrors.Name = "dataGridViewErrors";
            this.dataGridViewErrors.RowHeadersVisible = false;
            this.dataGridViewErrors.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewErrors.Size = new System.Drawing.Size(487, 219);
            this.dataGridViewErrors.TabIndex = 4;
            this.dataGridViewErrors.Visible = false;
            this.dataGridViewErrors.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewErrors_CellDoubleClick);
            // 
            // buttonShowDetails
            // 
            this.buttonShowDetails.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonShowDetails.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonShowDetails.ForeColor = System.Drawing.Color.Blue;
            this.buttonShowDetails.Image = ((System.Drawing.Image)(resources.GetObject("buttonShowDetails.Image")));
            this.buttonShowDetails.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonShowDetails.Location = new System.Drawing.Point(24, 132);
            this.buttonShowDetails.Name = "buttonShowDetails";
            this.buttonShowDetails.Size = new System.Drawing.Size(217, 29);
            this.buttonShowDetails.TabIndex = 8;
            this.buttonShowDetails.Text = "Show detailed error informations";
            this.buttonShowDetails.UseVisualStyleBackColor = true;
            this.buttonShowDetails.Click += new System.EventHandler(this.buttonShowDetails_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonClose.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.ForeColor = System.Drawing.Color.Blue;
            this.buttonClose.Image = ((System.Drawing.Image)(resources.GetObject("buttonClose.Image")));
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.Location = new System.Drawing.Point(335, 132);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(176, 29);
            this.buttonClose.TabIndex = 7;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // colMessage
            // 
            this.colMessage.DataPropertyName = "Message";
            this.colMessage.HeaderText = "Message";
            this.colMessage.Name = "colMessage";
            this.colMessage.Width = 150;
            // 
            // colType
            // 
            this.colType.DataPropertyName = "Type";
            this.colType.HeaderText = "Type";
            this.colType.Name = "colType";
            // 
            // colSource
            // 
            this.colSource.DataPropertyName = "Source";
            this.colSource.HeaderText = "Source";
            this.colSource.Name = "colSource";
            this.colSource.Width = 234;
            // 
            // ErrorDialog
            // 
            this.AcceptButton = this.buttonClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonClose;
            this.ClientSize = new System.Drawing.Size(534, 475);
            this.Controls.Add(this.buttonShowDetails);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonClipboardCopy);
            this.Controls.Add(this.dataGridViewErrors);
            this.Controls.Add(this.labelErrorMessage);
            this.Controls.Add(this.panelHeader);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MinimumSize = new System.Drawing.Size(540, 200);
            this.Name = "ErrorDialog";
            this.Text = "Error";
            this.panelHeader.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewErrors)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label labelErrorHeader;
        private System.Windows.Forms.PictureBox pictureBoxHeader;
        private System.Windows.Forms.Label labelErrorMessage;
        private System.Windows.Forms.Button buttonClipboardCopy;
        private System.Windows.Forms.DataGridView dataGridViewErrors;
        private System.Windows.Forms.Button buttonShowDetails;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMessage;
        private System.Windows.Forms.DataGridViewTextBoxColumn colType;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSource;
    }
}