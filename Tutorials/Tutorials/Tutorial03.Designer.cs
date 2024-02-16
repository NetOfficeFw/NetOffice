namespace TutorialsCS4
{
    partial class Tutorial03
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
            this.buttonDisposeChildInstances = new System.Windows.Forms.Button();
            this.buttonAddins = new System.Windows.Forms.Button();
            this.buttonWorkbook = new System.Windows.Forms.Button();
            this.buttonExcel = new System.Windows.Forms.Button();
            this.labelProxyCount = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.instanceMonitor1 = new NetOffice.Contribution.Controls.InstanceMonitor();
            this.SuspendLayout();
            // 
            // buttonDisposeChildInstances
            // 
            this.buttonDisposeChildInstances.Enabled = false;
            this.buttonDisposeChildInstances.Location = new System.Drawing.Point(476, 31);
            this.buttonDisposeChildInstances.Margin = new System.Windows.Forms.Padding(4);
            this.buttonDisposeChildInstances.Name = "buttonDisposeChildInstances";
            this.buttonDisposeChildInstances.Size = new System.Drawing.Size(263, 31);
            this.buttonDisposeChildInstances.TabIndex = 20;
            this.buttonDisposeChildInstances.Text = "Dispose Application Child Instances";
            this.buttonDisposeChildInstances.UseVisualStyleBackColor = true;
            this.buttonDisposeChildInstances.Click += new System.EventHandler(this.buttonDisposeChildInstances_Click);
            // 
            // buttonAddins
            // 
            this.buttonAddins.Enabled = false;
            this.buttonAddins.Location = new System.Drawing.Point(331, 31);
            this.buttonAddins.Margin = new System.Windows.Forms.Padding(4);
            this.buttonAddins.Name = "buttonAddins";
            this.buttonAddins.Size = new System.Drawing.Size(137, 31);
            this.buttonAddins.TabIndex = 19;
            this.buttonAddins.Text = "Enum Addins";
            this.buttonAddins.UseVisualStyleBackColor = true;
            this.buttonAddins.Click += new System.EventHandler(this.buttonAddins_Click);
            // 
            // buttonWorkbook
            // 
            this.buttonWorkbook.Enabled = false;
            this.buttonWorkbook.Location = new System.Drawing.Point(187, 31);
            this.buttonWorkbook.Margin = new System.Windows.Forms.Padding(4);
            this.buttonWorkbook.Name = "buttonWorkbook";
            this.buttonWorkbook.Size = new System.Drawing.Size(137, 31);
            this.buttonWorkbook.TabIndex = 18;
            this.buttonWorkbook.Text = "Add Workbook";
            this.buttonWorkbook.UseVisualStyleBackColor = true;
            this.buttonWorkbook.Click += new System.EventHandler(this.buttonWorkbook_Click);
            // 
            // buttonExcel
            // 
            this.buttonExcel.Location = new System.Drawing.Point(43, 31);
            this.buttonExcel.Margin = new System.Windows.Forms.Padding(4);
            this.buttonExcel.Name = "buttonExcel";
            this.buttonExcel.Size = new System.Drawing.Size(137, 31);
            this.buttonExcel.TabIndex = 17;
            this.buttonExcel.Text = "Start Excel";
            this.buttonExcel.UseVisualStyleBackColor = true;
            this.buttonExcel.Click += new System.EventHandler(this.buttonExcel_Click);
            // 
            // labelProxyCount
            // 
            this.labelProxyCount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelProxyCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProxyCount.Location = new System.Drawing.Point(247, 85);
            this.labelProxyCount.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelProxyCount.Name = "labelProxyCount";
            this.labelProxyCount.Size = new System.Drawing.Size(62, 24);
            this.labelProxyCount.TabIndex = 16;
            this.labelProxyCount.Text = "0";
            this.labelProxyCount.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(45, 87);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(197, 20);
            this.label1.TabIndex = 15;
            this.label1.Text = "Current COM Proxies open";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label2.Location = new System.Drawing.Point(45, 143);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 20);
            this.label2.TabIndex = 22;
            this.label2.Text = "Proxy Tree";
            // 
            // instanceMonitor1
            // 
            this.instanceMonitor1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.instanceMonitor1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.instanceMonitor1.Location = new System.Drawing.Point(43, 172);
            this.instanceMonitor1.Margin = new System.Windows.Forms.Padding(5);
            this.instanceMonitor1.Name = "instanceMonitor1";
            this.instanceMonitor1.Size = new System.Drawing.Size(696, 378);
            this.instanceMonitor1.TabIndex = 21;
            // 
            // Tutorial03
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.Controls.Add(this.label2);
            this.Controls.Add(this.instanceMonitor1);
            this.Controls.Add(this.buttonDisposeChildInstances);
            this.Controls.Add(this.buttonAddins);
            this.Controls.Add(this.buttonWorkbook);
            this.Controls.Add(this.buttonExcel);
            this.Controls.Add(this.labelProxyCount);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Tutorial03";
            this.Size = new System.Drawing.Size(800, 600);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonDisposeChildInstances;
        private System.Windows.Forms.Button buttonAddins;
        private System.Windows.Forms.Button buttonWorkbook;
        private System.Windows.Forms.Button buttonExcel;
        private System.Windows.Forms.Label labelProxyCount;
        private System.Windows.Forms.Label label1;
        private NetOffice.Contribution.Controls.InstanceMonitor instanceMonitor1;
        private System.Windows.Forms.Label label2;
    }
}
