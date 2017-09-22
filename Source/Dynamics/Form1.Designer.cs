namespace WindowsFormsApplication1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.StartExcelButton = new System.Windows.Forms.Button();
            this.QuitExcelButton = new System.Windows.Forms.Button();
            this.AddWorkbookButton = new System.Windows.Forms.Button();
            this.DisposeChildsButton = new System.Windows.Forms.Button();
            this.LogBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(585, 66);
            this.panel1.TabIndex = 0;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(16, 16);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(32, 32);
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.DarkGray;
            this.label2.Location = new System.Drawing.Point(58, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(268, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "This assembly contains only a reference to NetOffice.dll";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label1.Location = new System.Drawing.Point(56, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(277, 21);
            this.label1.TabIndex = 0;
            this.label1.Text = "Deal with unknown proxies at runtime.";
            // 
            // StartExcelButton
            // 
            this.StartExcelButton.Location = new System.Drawing.Point(16, 82);
            this.StartExcelButton.Name = "StartExcelButton";
            this.StartExcelButton.Size = new System.Drawing.Size(206, 23);
            this.StartExcelButton.TabIndex = 1;
            this.StartExcelButton.Text = "Start Invisible Excel";
            this.StartExcelButton.UseVisualStyleBackColor = true;
            this.StartExcelButton.Click += new System.EventHandler(this.StartExcelButton_Click);
            // 
            // QuitExcelButton
            // 
            this.QuitExcelButton.Location = new System.Drawing.Point(17, 210);
            this.QuitExcelButton.Name = "QuitExcelButton";
            this.QuitExcelButton.Size = new System.Drawing.Size(206, 23);
            this.QuitExcelButton.TabIndex = 2;
            this.QuitExcelButton.Text = "Quit Excel";
            this.QuitExcelButton.UseVisualStyleBackColor = true;
            this.QuitExcelButton.Click += new System.EventHandler(this.QuitExcelButton_Click);
            // 
            // AddWorkbookButton
            // 
            this.AddWorkbookButton.Location = new System.Drawing.Point(17, 126);
            this.AddWorkbookButton.Name = "AddWorkbookButton";
            this.AddWorkbookButton.Size = new System.Drawing.Size(206, 23);
            this.AddWorkbookButton.TabIndex = 3;
            this.AddWorkbookButton.Text = "Add 1 Workbook && Get Sheet Names";
            this.AddWorkbookButton.UseVisualStyleBackColor = true;
            this.AddWorkbookButton.Click += new System.EventHandler(this.AddWorkbookButton_Click);
            // 
            // DisposeChildsButton
            // 
            this.DisposeChildsButton.Location = new System.Drawing.Point(17, 168);
            this.DisposeChildsButton.Name = "DisposeChildsButton";
            this.DisposeChildsButton.Size = new System.Drawing.Size(206, 23);
            this.DisposeChildsButton.TabIndex = 4;
            this.DisposeChildsButton.Text = "Dispose All Excel Child Proxies";
            this.DisposeChildsButton.UseVisualStyleBackColor = true;
            this.DisposeChildsButton.Click += new System.EventHandler(this.DisposeChildsButton_Click);
            // 
            // LogBox
            // 
            this.LogBox.Location = new System.Drawing.Point(245, 98);
            this.LogBox.Multiline = true;
            this.LogBox.Name = "LogBox";
            this.LogBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.LogBox.Size = new System.Drawing.Size(313, 135);
            this.LogBox.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label3.Location = new System.Drawing.Point(245, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Action Log";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(585, 256);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.LogBox);
            this.Controls.Add(this.DisposeChildsButton);
            this.Controls.Add(this.AddWorkbookButton);
            this.Controls.Add(this.QuitExcelButton);
            this.Controls.Add(this.StartExcelButton);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "Introducing COMDynamicObject in C#";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button StartExcelButton;
        private System.Windows.Forms.Button QuitExcelButton;
        private System.Windows.Forms.Button AddWorkbookButton;
        private System.Windows.Forms.Button DisposeChildsButton;
        private System.Windows.Forms.TextBox LogBox;
        private System.Windows.Forms.Label label3;
    }
}

