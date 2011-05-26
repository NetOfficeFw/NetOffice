namespace Tutorial03
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
            this.panel2 = new System.Windows.Forms.Panel();
            this.linkTeqFaqEnglish = new System.Windows.Forms.LinkLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.linkTecDocEnglish = new System.Windows.Forms.LinkLabel();
            this.linkDocEnglish = new System.Windows.Forms.LinkLabel();
            this.linkFaqEnglish = new System.Windows.Forms.LinkLabel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.linkTeqFaqGerman = new System.Windows.Forms.LinkLabel();
            this.label2 = new System.Windows.Forms.Label();
            this.linkTecDocGerman = new System.Windows.Forms.LinkLabel();
            this.linkDocGerman = new System.Windows.Forms.LinkLabel();
            this.linkFaqGerman = new System.Windows.Forms.LinkLabel();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonStartExample
            // 
            this.buttonStartExample.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStartExample.Location = new System.Drawing.Point(25, 25);
            this.buttonStartExample.Name = "buttonStartExample";
            this.buttonStartExample.Size = new System.Drawing.Size(781, 30);
            this.buttonStartExample.TabIndex = 7;
            this.buttonStartExample.Text = "Start example";
            this.buttonStartExample.UseVisualStyleBackColor = true;
            this.buttonStartExample.Click += new System.EventHandler(this.buttonStartExample_Click);
            // 
            // richTextBoxInfo
            // 
            this.richTextBoxInfo.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.richTextBoxInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBoxInfo.Location = new System.Drawing.Point(25, 70);
            this.richTextBoxInfo.Name = "richTextBoxInfo";
            this.richTextBoxInfo.Size = new System.Drawing.Size(781, 227);
            this.richTextBoxInfo.TabIndex = 9;
            this.richTextBoxInfo.Text = resources.GetString("richTextBoxInfo.Text");
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.Control;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.linkTeqFaqEnglish);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.linkTecDocEnglish);
            this.panel2.Controls.Add(this.linkDocEnglish);
            this.panel2.Controls.Add(this.linkFaqEnglish);
            this.panel2.Location = new System.Drawing.Point(817, 25);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(154, 133);
            this.panel2.TabIndex = 24;
            // 
            // linkTeqFaqEnglish
            // 
            this.linkTeqFaqEnglish.AutoSize = true;
            this.linkTeqFaqEnglish.Location = new System.Drawing.Point(4, 103);
            this.linkTeqFaqEnglish.Name = "linkTeqFaqEnglish";
            this.linkTeqFaqEnglish.Size = new System.Drawing.Size(78, 13);
            this.linkTeqFaqEnglish.TabIndex = 14;
            this.linkTeqFaqEnglish.TabStop = true;
            this.linkTeqFaqEnglish.Tag = "http://netoffice.codeplex.com/wikipage?title=Tec_Faq_English";
            this.linkTeqFaqEnglish.Text = "Technical FAQ";
            this.linkTeqFaqEnglish.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "English";
            // 
            // linkTecDocEnglish
            // 
            this.linkTecDocEnglish.AutoSize = true;
            this.linkTecDocEnglish.Location = new System.Drawing.Point(3, 80);
            this.linkTecDocEnglish.Name = "linkTecDocEnglish";
            this.linkTecDocEnglish.Size = new System.Drawing.Size(129, 13);
            this.linkTecDocEnglish.TabIndex = 13;
            this.linkTecDocEnglish.TabStop = true;
            this.linkTecDocEnglish.Tag = "http://netoffice.codeplex.com/wikipage?title=Tec_Documentation_English";
            this.linkTecDocEnglish.Text = "Technical Documentation";
            this.linkTecDocEnglish.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // linkDocEnglish
            // 
            this.linkDocEnglish.AutoSize = true;
            this.linkDocEnglish.Location = new System.Drawing.Point(3, 34);
            this.linkDocEnglish.Name = "linkDocEnglish";
            this.linkDocEnglish.Size = new System.Drawing.Size(79, 13);
            this.linkDocEnglish.TabIndex = 15;
            this.linkDocEnglish.TabStop = true;
            this.linkDocEnglish.Tag = "http://netoffice.codeplex.com/documentation";
            this.linkDocEnglish.Text = "Documentation";
            this.linkDocEnglish.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // linkFaqEnglish
            // 
            this.linkFaqEnglish.AutoSize = true;
            this.linkFaqEnglish.Location = new System.Drawing.Point(3, 58);
            this.linkFaqEnglish.Name = "linkFaqEnglish";
            this.linkFaqEnglish.Size = new System.Drawing.Size(28, 13);
            this.linkFaqEnglish.TabIndex = 16;
            this.linkFaqEnglish.TabStop = true;
            this.linkFaqEnglish.Tag = "http://netoffice.codeplex.com/wikipage?title=FAQ";
            this.linkFaqEnglish.Text = "FAQ";
            this.linkFaqEnglish.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.linkTeqFaqGerman);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.linkTecDocGerman);
            this.panel1.Controls.Add(this.linkDocGerman);
            this.panel1.Controls.Add(this.linkFaqGerman);
            this.panel1.Location = new System.Drawing.Point(817, 164);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(154, 133);
            this.panel1.TabIndex = 23;
            // 
            // linkTeqFaqGerman
            // 
            this.linkTeqFaqGerman.AutoSize = true;
            this.linkTeqFaqGerman.Location = new System.Drawing.Point(4, 104);
            this.linkTeqFaqGerman.Name = "linkTeqFaqGerman";
            this.linkTeqFaqGerman.Size = new System.Drawing.Size(78, 13);
            this.linkTeqFaqGerman.TabIndex = 14;
            this.linkTeqFaqGerman.TabStop = true;
            this.linkTeqFaqGerman.Tag = "http://netoffice.codeplex.com/wikipage?title=Tec_Faq_German";
            this.linkTeqFaqGerman.Text = "Technical FAQ";
            this.linkTeqFaqGerman.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "German";
            // 
            // linkTecDocGerman
            // 
            this.linkTecDocGerman.AutoSize = true;
            this.linkTecDocGerman.Location = new System.Drawing.Point(3, 80);
            this.linkTecDocGerman.Name = "linkTecDocGerman";
            this.linkTecDocGerman.Size = new System.Drawing.Size(129, 13);
            this.linkTecDocGerman.TabIndex = 13;
            this.linkTecDocGerman.TabStop = true;
            this.linkTecDocGerman.Tag = "http://netoffice.codeplex.com/wikipage?title=Tec_Documentation_German";
            this.linkTecDocGerman.Text = "Technical Documentation";
            this.linkTecDocGerman.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // linkDocGerman
            // 
            this.linkDocGerman.AutoSize = true;
            this.linkDocGerman.Location = new System.Drawing.Point(3, 34);
            this.linkDocGerman.Name = "linkDocGerman";
            this.linkDocGerman.Size = new System.Drawing.Size(79, 13);
            this.linkDocGerman.TabIndex = 15;
            this.linkDocGerman.TabStop = true;
            this.linkDocGerman.Tag = "http://netoffice.codeplex.com/wikipage?title=Documentation_German";
            this.linkDocGerman.Text = "Documentation";
            this.linkDocGerman.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // linkFaqGerman
            // 
            this.linkFaqGerman.AutoSize = true;
            this.linkFaqGerman.Location = new System.Drawing.Point(3, 58);
            this.linkFaqGerman.Name = "linkFaqGerman";
            this.linkFaqGerman.Size = new System.Drawing.Size(28, 13);
            this.linkFaqGerman.TabIndex = 16;
            this.linkFaqGerman.TabStop = true;
            this.linkFaqGerman.Tag = "http://netoffice.codeplex.com/wikipage?title=FAQ_German";
            this.linkFaqGerman.Text = "FAQ";
            this.linkFaqGerman.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_LinkClicked);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(990, 321);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.richTextBoxInfo);
            this.Controls.Add(this.buttonStartExample);
            this.Name = "Form1";
            this.Text = "Tutorial03 - Using Dispose with event exporting Objects";
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonStartExample;
        private System.Windows.Forms.RichTextBox richTextBoxInfo;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.LinkLabel linkTeqFaqEnglish;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.LinkLabel linkTecDocEnglish;
        private System.Windows.Forms.LinkLabel linkDocEnglish;
        private System.Windows.Forms.LinkLabel linkFaqEnglish;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.LinkLabel linkTeqFaqGerman;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.LinkLabel linkTecDocGerman;
        private System.Windows.Forms.LinkLabel linkDocGerman;
        private System.Windows.Forms.LinkLabel linkFaqGerman;
    }
}

