namespace NetOffice.DeveloperToolbox.Controls.Error
{
    partial class ErrorControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorControl));
            this.linkLabelDiscussionBoard = new System.Windows.Forms.LinkLabel();
            this.listViewTrace = new System.Windows.Forms.ListView();
            this.columnHeaderSpace = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderMessage = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderSource = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.buttonCopyToClipboard = new System.Windows.Forms.Button();
            this.labelExitMessage = new System.Windows.Forms.Label();
            this.labelErrorMessage = new System.Windows.Forms.Label();
            this.labelErrorCaption = new System.Windows.Forms.Label();
            this.buttonOK = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // linkLabelDiscussionBoard
            // 
            this.linkLabelDiscussionBoard.AutoSize = true;
            this.linkLabelDiscussionBoard.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelDiscussionBoard.Location = new System.Drawing.Point(44, 132);
            this.linkLabelDiscussionBoard.Name = "linkLabelDiscussionBoard";
            this.linkLabelDiscussionBoard.Size = new System.Drawing.Size(148, 13);
            this.linkLabelDiscussionBoard.TabIndex = 35;
            this.linkLabelDiscussionBoard.TabStop = true;
            this.linkLabelDiscussionBoard.Tag = "http://netoffice.codeplex.com/discussions";
            this.linkLabelDiscussionBoard.Text = "NetOffice Discussion Board";
            this.linkLabelDiscussionBoard.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelDiscussionBoard_LinkClicked);
            // 
            // listViewTrace
            // 
            this.listViewTrace.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewTrace.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listViewTrace.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderSpace,
            this.columnHeaderMessage,
            this.columnHeaderType,
            this.columnHeaderSource});
            this.listViewTrace.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listViewTrace.ForeColor = System.Drawing.Color.Black;
            this.listViewTrace.FullRowSelect = true;
            this.listViewTrace.GridLines = true;
            this.listViewTrace.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listViewTrace.Location = new System.Drawing.Point(39, 189);
            this.listViewTrace.Name = "listViewTrace";
            this.listViewTrace.Size = new System.Drawing.Size(562, 245);
            this.listViewTrace.TabIndex = 34;
            this.listViewTrace.UseCompatibleStateImageBehavior = false;
            this.listViewTrace.View = System.Windows.Forms.View.Details;
            this.listViewTrace.Click += new System.EventHandler(this.listViewTrace_DoubleClick);
            // 
            // columnHeaderSpace
            // 
            this.columnHeaderSpace.Text = "";
            this.columnHeaderSpace.Width = 25;
            // 
            // columnHeaderMessage
            // 
            this.columnHeaderMessage.Text = "Message";
            this.columnHeaderMessage.Width = 246;
            // 
            // columnHeaderType
            // 
            this.columnHeaderType.Text = "Type";
            this.columnHeaderType.Width = 112;
            // 
            // columnHeaderSource
            // 
            this.columnHeaderSource.Text = "Source";
            this.columnHeaderSource.Width = 151;
            // 
            // buttonCopyToClipboard
            // 
            this.buttonCopyToClipboard.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCopyToClipboard.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonCopyToClipboard.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonCopyToClipboard.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.buttonCopyToClipboard.ForeColor = System.Drawing.Color.Blue;
            this.buttonCopyToClipboard.Location = new System.Drawing.Point(39, 440);
            this.buttonCopyToClipboard.Name = "buttonCopyToClipboard";
            this.buttonCopyToClipboard.Size = new System.Drawing.Size(562, 26);
            this.buttonCopyToClipboard.TabIndex = 33;
            this.buttonCopyToClipboard.Text = "Copy to Clipboard";
            this.buttonCopyToClipboard.UseVisualStyleBackColor = true;
            this.buttonCopyToClipboard.Click += new System.EventHandler(this.buttonCopyToClipboard_Click);
            // 
            // labelExitMessage
            // 
            this.labelExitMessage.AutoSize = true;
            this.labelExitMessage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelExitMessage.ForeColor = System.Drawing.Color.Black;
            this.labelExitMessage.Location = new System.Drawing.Point(98, 79);
            this.labelExitMessage.Name = "labelExitMessage";
            this.labelExitMessage.Size = new System.Drawing.Size(170, 13);
            this.labelExitMessage.TabIndex = 32;
            this.labelExitMessage.Text = "The application want close now";
            this.labelExitMessage.Visible = false;
            // 
            // labelErrorMessage
            // 
            this.labelErrorMessage.AutoSize = true;
            this.labelErrorMessage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorMessage.ForeColor = System.Drawing.Color.Black;
            this.labelErrorMessage.Location = new System.Drawing.Point(98, 53);
            this.labelErrorMessage.Name = "labelErrorMessage";
            this.labelErrorMessage.Size = new System.Drawing.Size(54, 13);
            this.labelErrorMessage.TabIndex = 31;
            this.labelErrorMessage.Text = "<Empty>";
            this.labelErrorMessage.Visible = false;
            // 
            // labelErrorCaption
            // 
            this.labelErrorCaption.AutoSize = true;
            this.labelErrorCaption.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorCaption.ForeColor = System.Drawing.Color.Black;
            this.labelErrorCaption.Location = new System.Drawing.Point(96, 23);
            this.labelErrorCaption.Name = "labelErrorCaption";
            this.labelErrorCaption.Size = new System.Drawing.Size(157, 21);
            this.labelErrorCaption.TabIndex = 30;
            this.labelErrorCaption.Text = "An error is occured.";
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonOK.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.buttonOK.ForeColor = System.Drawing.Color.Blue;
            this.buttonOK.Location = new System.Drawing.Point(508, 127);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(93, 25);
            this.buttonOK.TabIndex = 29;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(39, 21);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(37, 32);
            this.pictureBox1.TabIndex = 27;
            this.pictureBox1.TabStop = false;
            // 
            // ErrorControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.linkLabelDiscussionBoard);
            this.Controls.Add(this.listViewTrace);
            this.Controls.Add(this.buttonCopyToClipboard);
            this.Controls.Add(this.labelExitMessage);
            this.Controls.Add(this.labelErrorMessage);
            this.Controls.Add(this.labelErrorCaption);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.pictureBox1);
            this.Name = "ErrorControl";
            this.Size = new System.Drawing.Size(640, 514);
            this.Resize += new System.EventHandler(this.ErrorControl_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.LinkLabel linkLabelDiscussionBoard;
        private System.Windows.Forms.ListView listViewTrace;
        private System.Windows.Forms.ColumnHeader columnHeaderSpace;
        private System.Windows.Forms.ColumnHeader columnHeaderMessage;
        private System.Windows.Forms.ColumnHeader columnHeaderType;
        private System.Windows.Forms.ColumnHeader columnHeaderSource;
        private System.Windows.Forms.Button buttonCopyToClipboard;
        private System.Windows.Forms.Label labelExitMessage;
        private System.Windows.Forms.Label labelErrorMessage;
        private System.Windows.Forms.Label labelErrorCaption;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}
