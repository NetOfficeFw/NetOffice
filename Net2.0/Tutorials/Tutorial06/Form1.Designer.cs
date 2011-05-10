namespace Tutorial06
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
            this.buttonStartExample = new System.Windows.Forms.Button();
            this.richTextBoxInfo = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // buttonStartExample
            // 
            this.buttonStartExample.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStartExample.Location = new System.Drawing.Point(29, 26);
            this.buttonStartExample.Name = "buttonStartExample";
            this.buttonStartExample.Size = new System.Drawing.Size(801, 30);
            this.buttonStartExample.TabIndex = 9;
            this.buttonStartExample.Text = "Start example";
            this.buttonStartExample.UseVisualStyleBackColor = true;
            this.buttonStartExample.Click += new System.EventHandler(this.buttonStartExample_Click);
            // 
            // richTextBoxInfo
            // 
            this.richTextBoxInfo.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.richTextBoxInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.richTextBoxInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBoxInfo.Location = new System.Drawing.Point(29, 71);
            this.richTextBoxInfo.Name = "richTextBoxInfo";
            this.richTextBoxInfo.Size = new System.Drawing.Size(801, 237);
            this.richTextBoxInfo.TabIndex = 10;
            this.richTextBoxInfo.Text = resources.GetString("richTextBoxInfo.Text");
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(854, 329);
            this.Controls.Add(this.richTextBoxInfo);
            this.Controls.Add(this.buttonStartExample);
            this.Name = "Form1";
            this.Text = "Tutorial06 - Understand Variant Type";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonStartExample;
        private System.Windows.Forms.RichTextBox richTextBoxInfo;
    }
}

