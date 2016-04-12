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
            this.buttonTest = new System.Windows.Forms.Button();
            this.listViewResults = new System.Windows.Forms.ListView();
            this.columnHeaderIcon = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderDetails = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.labelCurrentTest = new System.Windows.Forms.Label();
            this.labelAdmin = new System.Windows.Forms.Label();
            this.labelRunningInstances = new System.Windows.Forms.Label();
            this.checkBoxExcel = new System.Windows.Forms.CheckBox();
            this.checkBoxWord = new System.Windows.Forms.CheckBox();
            this.checkBoxOutlook = new System.Windows.Forms.CheckBox();
            this.checkBoxPowerPoint = new System.Windows.Forms.CheckBox();
            this.checkBoxAccess = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBoxProject = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // buttonTest
            // 
            this.buttonTest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonTest.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonTest.Location = new System.Drawing.Point(507, 381);
            this.buttonTest.Name = "buttonTest";
            this.buttonTest.Size = new System.Drawing.Size(145, 25);
            this.buttonTest.TabIndex = 3;
            this.buttonTest.Text = "Do Test!";
            this.buttonTest.UseVisualStyleBackColor = true;
            this.buttonTest.Click += new System.EventHandler(this.buttonTest_Click);
            // 
            // listViewResults
            // 
            this.listViewResults.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewResults.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listViewResults.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderIcon,
            this.columnHeaderDetails});
            this.listViewResults.FullRowSelect = true;
            this.listViewResults.GridLines = true;
            this.listViewResults.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listViewResults.LargeImageList = this.imageList1;
            this.listViewResults.Location = new System.Drawing.Point(30, 31);
            this.listViewResults.MultiSelect = false;
            this.listViewResults.Name = "listViewResults";
            this.listViewResults.Size = new System.Drawing.Size(622, 327);
            this.listViewResults.SmallImageList = this.imageList1;
            this.listViewResults.StateImageList = this.imageList1;
            this.listViewResults.TabIndex = 6;
            this.listViewResults.UseCompatibleStateImageBehavior = false;
            this.listViewResults.View = System.Windows.Forms.View.Details;
            this.listViewResults.DoubleClick += new System.EventHandler(this.listViewResults_DoubleClick);
            // 
            // columnHeaderIcon
            // 
            this.columnHeaderIcon.Text = "Result";
            this.columnHeaderIcon.Width = 140;
            // 
            // columnHeaderDetails
            // 
            this.columnHeaderDetails.Text = "Details";
            this.columnHeaderDetails.Width = 450;
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
            this.labelCurrentTest.Location = new System.Drawing.Point(33, 383);
            this.labelCurrentTest.Name = "labelCurrentTest";
            this.labelCurrentTest.Size = new System.Drawing.Size(0, 13);
            this.labelCurrentTest.TabIndex = 7;
            // 
            // labelAdmin
            // 
            this.labelAdmin.AutoSize = true;
            this.labelAdmin.Location = new System.Drawing.Point(32, 8);
            this.labelAdmin.Name = "labelAdmin";
            this.labelAdmin.Size = new System.Drawing.Size(75, 13);
            this.labelAdmin.TabIndex = 8;
            this.labelAdmin.Text = "IsAdministrator";
            // 
            // labelRunningInstances
            // 
            this.labelRunningInstances.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelRunningInstances.AutoSize = true;
            this.labelRunningInstances.Cursor = System.Windows.Forms.Cursors.Hand;
            this.labelRunningInstances.ForeColor = System.Drawing.Color.Red;
            this.labelRunningInstances.Location = new System.Drawing.Point(412, 8);
            this.labelRunningInstances.Name = "labelRunningInstances";
            this.labelRunningInstances.Size = new System.Drawing.Size(240, 13);
            this.labelRunningInstances.TabIndex = 9;
            this.labelRunningInstances.Text = "Warning: one or more office application is running";
            this.labelRunningInstances.Visible = false;
            this.labelRunningInstances.Click += new System.EventHandler(this.labelRunningInstances_Click);
            // 
            // checkBoxExcel
            // 
            this.checkBoxExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBoxExcel.AutoSize = true;
            this.checkBoxExcel.Checked = true;
            this.checkBoxExcel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxExcel.Location = new System.Drawing.Point(680, 60);
            this.checkBoxExcel.Name = "checkBoxExcel";
            this.checkBoxExcel.Size = new System.Drawing.Size(49, 17);
            this.checkBoxExcel.TabIndex = 10;
            this.checkBoxExcel.Text = "Excel";
            this.checkBoxExcel.UseVisualStyleBackColor = true;
            // 
            // checkBoxWord
            // 
            this.checkBoxWord.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBoxWord.AutoSize = true;
            this.checkBoxWord.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxWord.Location = new System.Drawing.Point(680, 83);
            this.checkBoxWord.Name = "checkBoxWord";
            this.checkBoxWord.Size = new System.Drawing.Size(49, 17);
            this.checkBoxWord.TabIndex = 11;
            this.checkBoxWord.Text = "Word";
            this.checkBoxWord.UseVisualStyleBackColor = true;
            // 
            // checkBoxOutlook
            // 
            this.checkBoxOutlook.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBoxOutlook.AutoSize = true;
            this.checkBoxOutlook.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxOutlook.Location = new System.Drawing.Point(680, 106);
            this.checkBoxOutlook.Name = "checkBoxOutlook";
            this.checkBoxOutlook.Size = new System.Drawing.Size(60, 17);
            this.checkBoxOutlook.TabIndex = 12;
            this.checkBoxOutlook.Text = "Outlook";
            this.checkBoxOutlook.UseVisualStyleBackColor = true;
            // 
            // checkBoxPowerPoint
            // 
            this.checkBoxPowerPoint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBoxPowerPoint.AutoSize = true;
            this.checkBoxPowerPoint.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxPowerPoint.Location = new System.Drawing.Point(680, 129);
            this.checkBoxPowerPoint.Name = "checkBoxPowerPoint";
            this.checkBoxPowerPoint.Size = new System.Drawing.Size(80, 17);
            this.checkBoxPowerPoint.TabIndex = 13;
            this.checkBoxPowerPoint.Text = "Power Point";
            this.checkBoxPowerPoint.UseVisualStyleBackColor = true;
            // 
            // checkBoxAccess
            // 
            this.checkBoxAccess.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBoxAccess.AutoSize = true;
            this.checkBoxAccess.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxAccess.Location = new System.Drawing.Point(680, 152);
            this.checkBoxAccess.Name = "checkBoxAccess";
            this.checkBoxAccess.Size = new System.Drawing.Size(58, 17);
            this.checkBoxAccess.TabIndex = 14;
            this.checkBoxAccess.Text = "Access";
            this.checkBoxAccess.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(674, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 15;
            this.label1.Text = "Choose tests";
            // 
            // checkBoxProject
            // 
            this.checkBoxProject.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBoxProject.AutoSize = true;
            this.checkBoxProject.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxProject.Location = new System.Drawing.Point(680, 176);
            this.checkBoxProject.Name = "checkBoxProject";
            this.checkBoxProject.Size = new System.Drawing.Size(56, 17);
            this.checkBoxProject.TabIndex = 16;
            this.checkBoxProject.Text = "Project";
            this.checkBoxProject.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(789, 422);
            this.Controls.Add(this.checkBoxProject);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkBoxAccess);
            this.Controls.Add(this.checkBoxPowerPoint);
            this.Controls.Add(this.checkBoxOutlook);
            this.Controls.Add(this.checkBoxWord);
            this.Controls.Add(this.checkBoxExcel);
            this.Controls.Add(this.labelRunningInstances);
            this.Controls.Add(this.labelAdmin);
            this.Controls.Add(this.labelCurrentTest);
            this.Controls.Add(this.listViewResults);
            this.Controls.Add(this.buttonTest);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NetOffice Host for test applications";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonTest;
        private System.Windows.Forms.ListView listViewResults;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ColumnHeader columnHeaderIcon;
        private System.Windows.Forms.ColumnHeader columnHeaderDetails;
        private System.Windows.Forms.Label labelCurrentTest;
        private System.Windows.Forms.Label labelAdmin;
        private System.Windows.Forms.Label labelRunningInstances;
        private System.Windows.Forms.CheckBox checkBoxExcel;
        private System.Windows.Forms.CheckBox checkBoxWord;
        private System.Windows.Forms.CheckBox checkBoxOutlook;
        private System.Windows.Forms.CheckBox checkBoxPowerPoint;
        private System.Windows.Forms.CheckBox checkBoxAccess;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBoxProject;
    }
}

