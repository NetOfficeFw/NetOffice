namespace NOBuildTools.SourceUpdater
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxSource = new System.Windows.Forms.TextBox();
            this.textBoxDest = new System.Windows.Forms.TextBox();
            this.buttonChooseSource = new System.Windows.Forms.Button();
            this.buttonChooseDest = new System.Windows.Forms.Button();
            this.buttonStart = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Source";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 88);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Dest";
            // 
            // textBoxSource
            // 
            this.textBoxSource.Location = new System.Drawing.Point(76, 59);
            this.textBoxSource.Name = "textBoxSource";
            this.textBoxSource.Size = new System.Drawing.Size(606, 20);
            this.textBoxSource.TabIndex = 2;
            this.textBoxSource.Text = "C:\\LateBindingApi\\LateBindingApi.CodeGenerator.WFApplication\\bin\\Debug\\NetOffice";
            // 
            // textBoxDest
            // 
            this.textBoxDest.Location = new System.Drawing.Point(76, 85);
            this.textBoxDest.Name = "textBoxDest";
            this.textBoxDest.Size = new System.Drawing.Size(606, 20);
            this.textBoxDest.TabIndex = 3;
            this.textBoxDest.Text = "C:\\NetOffice\\Source";
            // 
            // buttonChooseSource
            // 
            this.buttonChooseSource.Location = new System.Drawing.Point(688, 57);
            this.buttonChooseSource.Name = "buttonChooseSource";
            this.buttonChooseSource.Size = new System.Drawing.Size(44, 23);
            this.buttonChooseSource.TabIndex = 4;
            this.buttonChooseSource.Text = "...";
            this.buttonChooseSource.UseVisualStyleBackColor = true;
            this.buttonChooseSource.Click += new System.EventHandler(this.buttonChooseSource_Click);
            // 
            // buttonChooseDest
            // 
            this.buttonChooseDest.Location = new System.Drawing.Point(688, 88);
            this.buttonChooseDest.Name = "buttonChooseDest";
            this.buttonChooseDest.Size = new System.Drawing.Size(44, 23);
            this.buttonChooseDest.TabIndex = 5;
            this.buttonChooseDest.Text = "...";
            this.buttonChooseDest.UseVisualStyleBackColor = true;
            this.buttonChooseDest.Click += new System.EventHandler(this.buttonChooseDest_Click);
            // 
            // buttonStart
            // 
            this.buttonStart.Location = new System.Drawing.Point(596, 126);
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.Size = new System.Drawing.Size(77, 23);
            this.buttonStart.TabIndex = 6;
            this.buttonStart.Text = "Start";
            this.buttonStart.UseVisualStyleBackColor = true;
            this.buttonStart.Click += new System.EventHandler(this.buttonStart_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.Location = new System.Drawing.Point(76, 129);
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.Size = new System.Drawing.Size(485, 20);
            this.textBoxLog.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(73, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(624, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "This tool update the existing NetOffice source folder with a newer version. (repl" +
                "ace by windows explorer kills .svn meta informations)";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(761, 180);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.buttonStart);
            this.Controls.Add(this.buttonChooseDest);
            this.Controls.Add(this.buttonChooseSource);
            this.Controls.Add(this.textBoxDest);
            this.Controls.Add(this.textBoxSource);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "SVN friendly CodeIntegrator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxSource;
        private System.Windows.Forms.TextBox textBoxDest;
        private System.Windows.Forms.Button buttonChooseSource;
        private System.Windows.Forms.Button buttonChooseDest;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.Label label3;
    }
}

