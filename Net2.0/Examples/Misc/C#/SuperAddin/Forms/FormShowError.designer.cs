namespace SuperAddin
{
    partial class FormShowError
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormShowError));
            this.pictureBoxError = new System.Windows.Forms.PictureBox();
            this.labelErrorFooter = new System.Windows.Forms.Label();
            this.buttonDetails = new System.Windows.Forms.Button();
            this.buttonOk = new System.Windows.Forms.Button();
            this.listViewExceptions = new System.Windows.Forms.ListView();
            this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader4 = new System.Windows.Forms.ColumnHeader();
            this.labelErrorHeader = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxError)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBoxError
            // 
            this.pictureBoxError.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxError.Image")));
            this.pictureBoxError.Location = new System.Drawing.Point(37, 25);
            this.pictureBoxError.Name = "pictureBoxError";
            this.pictureBoxError.Size = new System.Drawing.Size(51, 47);
            this.pictureBoxError.TabIndex = 11;
            this.pictureBoxError.TabStop = false;
            // 
            // labelErrorFooter
            // 
            this.labelErrorFooter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelErrorFooter.BackColor = System.Drawing.SystemColors.Control;
            this.labelErrorFooter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorFooter.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.labelErrorFooter.Location = new System.Drawing.Point(96, 78);
            this.labelErrorFooter.Name = "labelErrorFooter";
            this.labelErrorFooter.Size = new System.Drawing.Size(393, 43);
            this.labelErrorFooter.TabIndex = 10;
            this.labelErrorFooter.Text = "labelErrorFooter";
            this.labelErrorFooter.Visible = false;
            // 
            // buttonDetails
            // 
            this.buttonDetails.Location = new System.Drawing.Point(27, 124);
            this.buttonDetails.Name = "buttonDetails";
            this.buttonDetails.Size = new System.Drawing.Size(87, 22);
            this.buttonDetails.TabIndex = 9;
            this.buttonDetails.Text = "<< Details";
            this.buttonDetails.UseVisualStyleBackColor = true;
            this.buttonDetails.Click += new System.EventHandler(this.buttonDetails_Click);
            // 
            // buttonOk
            // 
            this.buttonOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOk.Location = new System.Drawing.Point(402, 124);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(87, 22);
            this.buttonOk.TabIndex = 8;
            this.buttonOk.Text = "Ok";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // listViewExceptions
            // 
            this.listViewExceptions.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewExceptions.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader4});
            this.listViewExceptions.FullRowSelect = true;
            this.listViewExceptions.GridLines = true;
            this.listViewExceptions.HideSelection = false;
            this.listViewExceptions.Location = new System.Drawing.Point(24, 172);
            this.listViewExceptions.Name = "listViewExceptions";
            this.listViewExceptions.Size = new System.Drawing.Size(483, 177);
            this.listViewExceptions.TabIndex = 7;
            this.listViewExceptions.UseCompatibleStateImageBehavior = false;
            this.listViewExceptions.View = System.Windows.Forms.View.Details;
            this.listViewExceptions.Resize += new System.EventHandler(this.listViewExceptions_Resize);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Nr";
            this.columnHeader1.Width = 38;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Modul";
            this.columnHeader2.Width = 91;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Type";
            this.columnHeader3.Width = 83;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Text";
            this.columnHeader4.Width = 218;
            // 
            // labelErrorHeader
            // 
            this.labelErrorHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelErrorHeader.BackColor = System.Drawing.SystemColors.Control;
            this.labelErrorHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorHeader.ForeColor = System.Drawing.SystemColors.ControlText;
            this.labelErrorHeader.Location = new System.Drawing.Point(96, 24);
            this.labelErrorHeader.Name = "labelErrorHeader";
            this.labelErrorHeader.Size = new System.Drawing.Size(393, 39);
            this.labelErrorHeader.TabIndex = 6;
            this.labelErrorHeader.Text = "labelErrorHeader";
            this.labelErrorHeader.Visible = false;
            // 
            // FormShowError
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(531, 373);
            this.Controls.Add(this.pictureBoxError);
            this.Controls.Add(this.labelErrorFooter);
            this.Controls.Add(this.buttonDetails);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.listViewExceptions);
            this.Controls.Add(this.labelErrorHeader);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormShowError";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Error";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxError)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBoxError;
        private System.Windows.Forms.Label labelErrorFooter;
        private System.Windows.Forms.Button buttonDetails;
        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.ListView listViewExceptions;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.Label labelErrorHeader;

    }
}
