namespace OutlookExamplesCS4
{
    partial class Example03
    {
        /// <summary> 
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary> 
        /// Erforderliche Methode für die Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Example03));
            this.textBoxBody = new System.Windows.Forms.TextBox();
            this.textBoxSubject = new System.Windows.Forms.TextBox();
            this.textBoxReciever = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.buttonStartExample = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxBody
            // 
            this.textBoxBody.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxBody.Location = new System.Drawing.Point(91, 162);
            this.textBoxBody.Multiline = true;
            this.textBoxBody.Name = "textBoxBody";
            this.textBoxBody.Size = new System.Drawing.Size(610, 114);
            this.textBoxBody.TabIndex = 19;
            this.textBoxBody.Text = "This is a mail from NetOffice example.";
            this.textBoxBody.WordWrap = false;
            // 
            // textBoxSubject
            // 
            this.textBoxSubject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxSubject.Location = new System.Drawing.Point(91, 136);
            this.textBoxSubject.Name = "textBoxSubject";
            this.textBoxSubject.Size = new System.Drawing.Size(610, 20);
            this.textBoxSubject.TabIndex = 18;
            this.textBoxSubject.Text = "NetOffice Example Mail";
            // 
            // textBoxReciever
            // 
            this.textBoxReciever.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxReciever.Location = new System.Drawing.Point(91, 110);
            this.textBoxReciever.Name = "textBoxReciever";
            this.textBoxReciever.Size = new System.Drawing.Size(610, 20);
            this.textBoxReciever.TabIndex = 17;
            this.textBoxReciever.Text = "public.sebastian@web.de";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(36, 165);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 16;
            this.label3.Text = "Body";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(35, 140);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Subject";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(35, 113);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 13);
            this.label1.TabIndex = 14;
            this.label1.Text = "Reciever";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.textBox1.Location = new System.Drawing.Point(91, 68);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(610, 25);
            this.textBox1.TabIndex = 13;
            this.textBox1.Text = "this example shows you how to send a mail.";
            // 
            // buttonStartExample
            // 
            this.buttonStartExample.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonStartExample.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStartExample.Image = ((System.Drawing.Image)(resources.GetObject("buttonStartExample.Image")));
            this.buttonStartExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonStartExample.Location = new System.Drawing.Point(36, 22);
            this.buttonStartExample.Name = "buttonStartExample";
            this.buttonStartExample.Size = new System.Drawing.Size(665, 30);
            this.buttonStartExample.TabIndex = 12;
            this.buttonStartExample.Text = "Start example";
            this.buttonStartExample.UseVisualStyleBackColor = true;
            this.buttonStartExample.Click += new System.EventHandler(this.buttonStartExample_Click);
            // 
            // Example03
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.textBoxBody);
            this.Controls.Add(this.textBoxSubject);
            this.Controls.Add(this.textBoxReciever);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.buttonStartExample);
            this.Name = "Example03";
            this.Size = new System.Drawing.Size(739, 304);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxBody;
        private System.Windows.Forms.TextBox textBoxSubject;
        private System.Windows.Forms.TextBox textBoxReciever;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button buttonStartExample;
    }
}
