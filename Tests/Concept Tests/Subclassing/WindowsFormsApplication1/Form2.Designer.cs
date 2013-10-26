namespace WindowsFormsApplication1
{
    partial class Form2
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
            this.UserControlTest = new WindowsFormsApplication1.UserControl1();
            this.SuspendLayout();
            // 
            // UserControlTest
            // 
            this.UserControlTest.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.UserControlTest.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UserControlTest.Location = new System.Drawing.Point(0, 0);
            this.UserControlTest.Name = "UserControlTest";
            this.UserControlTest.Size = new System.Drawing.Size(489, 255);
            this.UserControlTest.TabIndex = 0;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(489, 255);
            this.Controls.Add(this.UserControlTest);
            this.Name = "Form2";
            this.Text = "Form2";
            this.ResumeLayout(false);

        }

        #endregion

        internal UserControl1 UserControlTest;

    }
}