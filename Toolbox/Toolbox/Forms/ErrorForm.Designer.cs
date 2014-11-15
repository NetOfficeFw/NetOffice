namespace NetOffice.DeveloperToolbox.Forms
{
    partial class ErrorForm
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorForm));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonDetails = new System.Windows.Forms.Button();
            this.labelErrorCaption = new System.Windows.Forms.Label();
            this.labelErrorMessage = new System.Windows.Forms.Label();
            this.labelExitMessage = new System.Windows.Forms.Label();
            this.buttonCopyToClipboard = new System.Windows.Forms.Button();
            this.listViewTrace = new System.Windows.Forms.ListView();
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader8 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.linkLabelDiscussionBoard = new System.Windows.Forms.LinkLabel();
            this.pictureBoxSplitter1 = new System.Windows.Forms.PictureBox();
            this.pictureBoxSplitter2 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSplitter1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSplitter2)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(37, 16);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(37, 32);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // buttonOK
            // 
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonOK.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOK.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonOK.Location = new System.Drawing.Point(506, 122);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(93, 25);
            this.buttonOK.TabIndex = 8;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonDetails
            // 
            this.buttonDetails.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonDetails.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonDetails.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonDetails.Location = new System.Drawing.Point(39, 121);
            this.buttonDetails.Name = "buttonDetails";
            this.buttonDetails.Size = new System.Drawing.Size(93, 26);
            this.buttonDetails.TabIndex = 7;
            this.buttonDetails.Text = "<< Details";
            this.buttonDetails.UseVisualStyleBackColor = true;
            this.buttonDetails.Click += new System.EventHandler(this.buttonDetails_Click);
            // 
            // labelErrorCaption
            // 
            this.labelErrorCaption.AutoSize = true;
            this.labelErrorCaption.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorCaption.Location = new System.Drawing.Point(96, 19);
            this.labelErrorCaption.Name = "labelErrorCaption";
            this.labelErrorCaption.Size = new System.Drawing.Size(256, 21);
            this.labelErrorCaption.TabIndex = 18;
            this.labelErrorCaption.Text = "Leider ist ein Fehler aufgetreten.";
            // 
            // labelErrorMessage
            // 
            this.labelErrorMessage.AutoSize = true;
            this.labelErrorMessage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelErrorMessage.Location = new System.Drawing.Point(96, 48);
            this.labelErrorMessage.Name = "labelErrorMessage";
            this.labelErrorMessage.Size = new System.Drawing.Size(44, 13);
            this.labelErrorMessage.TabIndex = 19;
            this.labelErrorMessage.Text = "<Leer>";
            this.labelErrorMessage.Visible = false;
            // 
            // labelExitMessage
            // 
            this.labelExitMessage.AutoSize = true;
            this.labelExitMessage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelExitMessage.Location = new System.Drawing.Point(96, 74);
            this.labelExitMessage.Name = "labelExitMessage";
            this.labelExitMessage.Size = new System.Drawing.Size(156, 13);
            this.labelExitMessage.TabIndex = 20;
            this.labelExitMessage.Text = "Das Programm wird beendet.";
            this.labelExitMessage.Visible = false;
            // 
            // buttonCopyToClipboard
            // 
            this.buttonCopyToClipboard.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonCopyToClipboard.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonCopyToClipboard.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonCopyToClipboard.Location = new System.Drawing.Point(37, 435);
            this.buttonCopyToClipboard.Name = "buttonCopyToClipboard";
            this.buttonCopyToClipboard.Size = new System.Drawing.Size(562, 26);
            this.buttonCopyToClipboard.TabIndex = 22;
            this.buttonCopyToClipboard.Text = "In die Zwischenablage kopieren";
            this.buttonCopyToClipboard.UseVisualStyleBackColor = true;
            this.buttonCopyToClipboard.Click += new System.EventHandler(this.buttonCopyToClipboard_Click);
            // 
            // listViewTrace
            // 
            this.listViewTrace.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listViewTrace.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader5,
            this.columnHeader6,
            this.columnHeader7,
            this.columnHeader8});
            this.listViewTrace.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listViewTrace.FullRowSelect = true;
            this.listViewTrace.GridLines = true;
            this.listViewTrace.Location = new System.Drawing.Point(37, 184);
            this.listViewTrace.Name = "listViewTrace";
            this.listViewTrace.Size = new System.Drawing.Size(562, 245);
            this.listViewTrace.TabIndex = 23;
            this.listViewTrace.UseCompatibleStateImageBehavior = false;
            this.listViewTrace.View = System.Windows.Forms.View.Details;
            this.listViewTrace.DoubleClick += new System.EventHandler(this.listViewTrace_DoubleClick);
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "";
            this.columnHeader5.Width = 25;
            // 
            // columnHeader6
            // 
            this.columnHeader6.Text = "Message";
            this.columnHeader6.Width = 246;
            // 
            // columnHeader7
            // 
            this.columnHeader7.Text = "Type";
            this.columnHeader7.Width = 112;
            // 
            // columnHeader8
            // 
            this.columnHeader8.Text = "Source";
            this.columnHeader8.Width = 151;
            // 
            // linkLabelDiscussionBoard
            // 
            this.linkLabelDiscussionBoard.AutoSize = true;
            this.linkLabelDiscussionBoard.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelDiscussionBoard.Location = new System.Drawing.Point(158, 127);
            this.linkLabelDiscussionBoard.Name = "linkLabelDiscussionBoard";
            this.linkLabelDiscussionBoard.Size = new System.Drawing.Size(148, 13);
            this.linkLabelDiscussionBoard.TabIndex = 24;
            this.linkLabelDiscussionBoard.TabStop = true;
            this.linkLabelDiscussionBoard.Tag = "http://netoffice.codeplex.com/discussions";
            this.linkLabelDiscussionBoard.Text = "NetOffice Discussion Board";
            this.linkLabelDiscussionBoard.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelDiscussionBoard_LinkClicked);
            // 
            // pictureBoxSplitter1
            // 
            this.pictureBoxSplitter1.BackColor = System.Drawing.Color.Red;
            this.pictureBoxSplitter1.Location = new System.Drawing.Point(39, 170);
            this.pictureBoxSplitter1.Name = "pictureBoxSplitter1";
            this.pictureBoxSplitter1.Size = new System.Drawing.Size(100, 10);
            this.pictureBoxSplitter1.TabIndex = 25;
            this.pictureBoxSplitter1.TabStop = false;
            this.pictureBoxSplitter1.Visible = false;
            // 
            // pictureBoxSplitter2
            // 
            this.pictureBoxSplitter2.BackColor = System.Drawing.Color.Red;
            this.pictureBoxSplitter2.Location = new System.Drawing.Point(37, 479);
            this.pictureBoxSplitter2.Name = "pictureBoxSplitter2";
            this.pictureBoxSplitter2.Size = new System.Drawing.Size(100, 10);
            this.pictureBoxSplitter2.TabIndex = 26;
            this.pictureBoxSplitter2.TabStop = false;
            this.pictureBoxSplitter2.Visible = false;
            // 
            // ErrorForm
            // 
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.CancelButton = this.buttonOK;
            this.ClientSize = new System.Drawing.Size(634, 489);
            this.Controls.Add(this.pictureBoxSplitter2);
            this.Controls.Add(this.pictureBoxSplitter1);
            this.Controls.Add(this.linkLabelDiscussionBoard);
            this.Controls.Add(this.listViewTrace);
            this.Controls.Add(this.buttonCopyToClipboard);
            this.Controls.Add(this.labelExitMessage);
            this.Controls.Add(this.labelErrorMessage);
            this.Controls.Add(this.labelErrorCaption);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonDetails);
            this.Controls.Add(this.pictureBox1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ErrorForm";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Fehler";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSplitter1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxSplitter2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonDetails;
        private System.Windows.Forms.Label labelErrorCaption;
        private System.Windows.Forms.Label labelErrorMessage;
        private System.Windows.Forms.Label labelExitMessage;
        private System.Windows.Forms.Button buttonCopyToClipboard;
        private System.Windows.Forms.ListView listViewTrace;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.ColumnHeader columnHeader6;
        private System.Windows.Forms.ColumnHeader columnHeader7;
        private System.Windows.Forms.ColumnHeader columnHeader8;
        private System.Windows.Forms.LinkLabel linkLabelDiscussionBoard;
        private System.Windows.Forms.PictureBox pictureBoxSplitter1;
        private System.Windows.Forms.PictureBox pictureBoxSplitter2;

    }
}
