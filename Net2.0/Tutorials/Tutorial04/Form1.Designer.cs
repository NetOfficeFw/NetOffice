namespace Tutorial04
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
            this.buttonAddRemoveWorkbook = new System.Windows.Forms.Button();
            this.buttonAddins = new System.Windows.Forms.Button();
            this.buttonWorkbook = new System.Windows.Forms.Button();
            this.buttonExcel = new System.Windows.Forms.Button();
            this.labelProxyCount = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.richTextBoxInfo = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // buttonAddRemoveWorkbook
            // 
            this.buttonAddRemoveWorkbook.Enabled = false;
            this.buttonAddRemoveWorkbook.Location = new System.Drawing.Point(14, 190);
            this.buttonAddRemoveWorkbook.Name = "buttonAddRemoveWorkbook";
            this.buttonAddRemoveWorkbook.Size = new System.Drawing.Size(103, 25);
            this.buttonAddRemoveWorkbook.TabIndex = 13;
            this.buttonAddRemoveWorkbook.Text = "Add && Remove";
            this.buttonAddRemoveWorkbook.UseVisualStyleBackColor = true;
            this.buttonAddRemoveWorkbook.Click += new System.EventHandler(this.buttonAddRemoveWorkbook_Click);
            // 
            // buttonAddins
            // 
            this.buttonAddins.Enabled = false;
            this.buttonAddins.Location = new System.Drawing.Point(14, 114);
            this.buttonAddins.Name = "buttonAddins";
            this.buttonAddins.Size = new System.Drawing.Size(103, 25);
            this.buttonAddins.TabIndex = 11;
            this.buttonAddins.Text = "Enum Addins";
            this.buttonAddins.UseVisualStyleBackColor = true;
            this.buttonAddins.Click += new System.EventHandler(this.buttonAddins_Click);
            // 
            // buttonWorkbook
            // 
            this.buttonWorkbook.Enabled = false;
            this.buttonWorkbook.Location = new System.Drawing.Point(14, 73);
            this.buttonWorkbook.Name = "buttonWorkbook";
            this.buttonWorkbook.Size = new System.Drawing.Size(103, 25);
            this.buttonWorkbook.TabIndex = 10;
            this.buttonWorkbook.Text = "Add Workbook";
            this.buttonWorkbook.UseVisualStyleBackColor = true;
            this.buttonWorkbook.Click += new System.EventHandler(this.buttonWorkbook_Click);
            // 
            // buttonExcel
            // 
            this.buttonExcel.Location = new System.Drawing.Point(14, 31);
            this.buttonExcel.Name = "buttonExcel";
            this.buttonExcel.Size = new System.Drawing.Size(103, 25);
            this.buttonExcel.TabIndex = 9;
            this.buttonExcel.Text = "Start Excel";
            this.buttonExcel.UseVisualStyleBackColor = true;
            this.buttonExcel.Click += new System.EventHandler(this.buttonExcel_Click);
            // 
            // labelProxyCount
            // 
            this.labelProxyCount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelProxyCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProxyCount.Location = new System.Drawing.Point(451, 360);
            this.labelProxyCount.Name = "labelProxyCount";
            this.labelProxyCount.Size = new System.Drawing.Size(47, 20);
            this.labelProxyCount.TabIndex = 8;
            this.labelProxyCount.Text = "0";
            this.labelProxyCount.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(149, 360);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(296, 20);
            this.label1.TabIndex = 7;
            this.label1.Text = "Current COM Objects open in application";
            // 
            // richTextBoxInfo
            // 
            this.richTextBoxInfo.Location = new System.Drawing.Point(153, 31);
            this.richTextBoxInfo.Name = "richTextBoxInfo";
            this.richTextBoxInfo.Size = new System.Drawing.Size(672, 312);
            this.richTextBoxInfo.TabIndex = 14;
            this.richTextBoxInfo.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(837, 399);
            this.Controls.Add(this.richTextBoxInfo);
            this.Controls.Add(this.buttonAddRemoveWorkbook);
            this.Controls.Add(this.buttonAddins);
            this.Controls.Add(this.buttonWorkbook);
            this.Controls.Add(this.buttonExcel);
            this.Controls.Add(this.labelProxyCount);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Tutorial04 - Observable COM Proxy Count";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonAddRemoveWorkbook;
        private System.Windows.Forms.Button buttonAddins;
        private System.Windows.Forms.Button buttonWorkbook;
        private System.Windows.Forms.Button buttonExcel;
        private System.Windows.Forms.Label labelProxyCount;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox richTextBoxInfo;
    }
}

