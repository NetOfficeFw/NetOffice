namespace Host
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.textBoxFolder = new System.Windows.Forms.TextBox();
            this.labelFolder = new System.Windows.Forms.Label();
            this.buttonTest = new System.Windows.Forms.Button();
            this.buttonSelectFolder = new System.Windows.Forms.Button();
            this.buttonOpenFolder = new System.Windows.Forms.Button();
            this.listViewResults = new System.Windows.Forms.ListView();
            this.columnHeaderIcon = new System.Windows.Forms.ColumnHeader();
            this.columnHeaderDetails = new System.Windows.Forms.ColumnHeader();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.labelCurrentTest = new System.Windows.Forms.Label();
            this.buttonSetDefaultFolder = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxFolder
            // 
            this.textBoxFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxFolder.Location = new System.Drawing.Point(53, 23);
            this.textBoxFolder.Name = "textBoxFolder";
            this.textBoxFolder.ReadOnly = true;
            this.textBoxFolder.Size = new System.Drawing.Size(488, 20);
            this.textBoxFolder.TabIndex = 1;
            // 
            // labelFolder
            // 
            this.labelFolder.AutoSize = true;
            this.labelFolder.Location = new System.Drawing.Point(12, 27);
            this.labelFolder.Name = "labelFolder";
            this.labelFolder.Size = new System.Drawing.Size(36, 13);
            this.labelFolder.TabIndex = 2;
            this.labelFolder.Text = "Folder";
            // 
            // buttonTest
            // 
            this.buttonTest.Location = new System.Drawing.Point(53, 49);
            this.buttonTest.Name = "buttonTest";
            this.buttonTest.Size = new System.Drawing.Size(145, 25);
            this.buttonTest.TabIndex = 3;
            this.buttonTest.Text = "Do Test!";
            this.buttonTest.UseVisualStyleBackColor = true;
            this.buttonTest.Click += new System.EventHandler(this.buttonTest_Click);
            // 
            // buttonSelectFolder
            // 
            this.buttonSelectFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSelectFolder.Location = new System.Drawing.Point(547, 20);
            this.buttonSelectFolder.Name = "buttonSelectFolder";
            this.buttonSelectFolder.Size = new System.Drawing.Size(41, 25);
            this.buttonSelectFolder.TabIndex = 4;
            this.buttonSelectFolder.Text = "...";
            this.buttonSelectFolder.UseVisualStyleBackColor = true;
            this.buttonSelectFolder.Click += new System.EventHandler(this.buttonSelectFolder_Click);
            // 
            // buttonOpenFolder
            // 
            this.buttonOpenFolder.Location = new System.Drawing.Point(376, 49);
            this.buttonOpenFolder.Name = "buttonOpenFolder";
            this.buttonOpenFolder.Size = new System.Drawing.Size(165, 25);
            this.buttonOpenFolder.TabIndex = 5;
            this.buttonOpenFolder.Text = "Open Folder";
            this.buttonOpenFolder.UseVisualStyleBackColor = true;
            this.buttonOpenFolder.Click += new System.EventHandler(this.buttonOpenFolder_Click);
            // 
            // listViewResults
            // 
            this.listViewResults.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewResults.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderIcon,
            this.columnHeaderDetails});
            this.listViewResults.LargeImageList = this.imageList1;
            this.listViewResults.Location = new System.Drawing.Point(53, 112);
            this.listViewResults.Name = "listViewResults";
            this.listViewResults.Size = new System.Drawing.Size(488, 212);
            this.listViewResults.SmallImageList = this.imageList1;
            this.listViewResults.StateImageList = this.imageList1;
            this.listViewResults.TabIndex = 6;
            this.listViewResults.UseCompatibleStateImageBehavior = false;
            this.listViewResults.View = System.Windows.Forms.View.Details;
            // 
            // columnHeaderIcon
            // 
            this.columnHeaderIcon.Text = "Result";
            this.columnHeaderIcon.Width = 140;
            // 
            // columnHeaderDetails
            // 
            this.columnHeaderDetails.Text = "Details";
            this.columnHeaderDetails.Width = 300;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "tick.png");
            this.imageList1.Images.SetKeyName(1, "exclamation.png");
            // 
            // labelCurrentTest
            // 
            this.labelCurrentTest.AutoSize = true;
            this.labelCurrentTest.Location = new System.Drawing.Point(60, 96);
            this.labelCurrentTest.Name = "labelCurrentTest";
            this.labelCurrentTest.Size = new System.Drawing.Size(0, 13);
            this.labelCurrentTest.TabIndex = 7;
            // 
            // buttonSetDefaultFolder
            // 
            this.buttonSetDefaultFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSetDefaultFolder.Location = new System.Drawing.Point(205, 49);
            this.buttonSetDefaultFolder.Name = "buttonSetDefaultFolder";
            this.buttonSetDefaultFolder.Size = new System.Drawing.Size(165, 25);
            this.buttonSetDefaultFolder.TabIndex = 8;
            this.buttonSetDefaultFolder.Text = "Set Default Folder";
            this.buttonSetDefaultFolder.UseVisualStyleBackColor = true;
            this.buttonSetDefaultFolder.Click += new System.EventHandler(this.buttonSetDefaultFolder_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 336);
            this.Controls.Add(this.buttonSetDefaultFolder);
            this.Controls.Add(this.labelCurrentTest);
            this.Controls.Add(this.listViewResults);
            this.Controls.Add(this.buttonOpenFolder);
            this.Controls.Add(this.buttonSelectFolder);
            this.Controls.Add(this.buttonTest);
            this.Controls.Add(this.labelFolder);
            this.Controls.Add(this.textBoxFolder);
            this.Name = "Form1";
            this.Text = "NetOffice Host for test applications";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxFolder;
        private System.Windows.Forms.Label labelFolder;
        private System.Windows.Forms.Button buttonTest;
        private System.Windows.Forms.Button buttonSelectFolder;
        private System.Windows.Forms.Button buttonOpenFolder;
        private System.Windows.Forms.ListView listViewResults;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ColumnHeader columnHeaderIcon;
        private System.Windows.Forms.ColumnHeader columnHeaderDetails;
        private System.Windows.Forms.Label labelCurrentTest;
        private System.Windows.Forms.Button buttonSetDefaultFolder;
    }
}

