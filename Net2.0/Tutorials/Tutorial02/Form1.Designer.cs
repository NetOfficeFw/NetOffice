namespace Tutorial02
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
            this.label1 = new System.Windows.Forms.Label();
            this.labelProxyCount = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(156, 352);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(296, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Current COM Objects open in application";
            // 
            // labelProxyCount
            // 
            this.labelProxyCount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelProxyCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProxyCount.Location = new System.Drawing.Point(458, 352);
            this.labelProxyCount.Name = "labelProxyCount";
            this.labelProxyCount.Size = new System.Drawing.Size(47, 20);
            this.labelProxyCount.TabIndex = 1;
            this.labelProxyCount.Text = "0";
            this.labelProxyCount.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(21, 23);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(103, 25);
            this.button1.TabIndex = 2;
            this.button1.Text = "Start Excel";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(21, 65);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(103, 25);
            this.button2.TabIndex = 3;
            this.button2.Text = "Add Workbook";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(21, 106);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(103, 25);
            this.button3.TabIndex = 4;
            this.button3.Text = "Enum Addins";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.textBox1.Location = new System.Drawing.Point(150, 23);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(670, 312);
            this.textBox1.TabIndex = 5;
            this.textBox1.Text = resources.GetString("textBox1.Text");
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(845, 389);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.labelProxyCount);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Tutorial02 - Observable COM Proxy Count";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelProxyCount;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox textBox1;
    }
}

