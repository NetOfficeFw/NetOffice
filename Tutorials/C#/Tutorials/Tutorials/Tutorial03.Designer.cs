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
            this.buttonAddRemoveWorkbook = new System.Windows.Forms.Button();
            this.buttonAddins = new System.Windows.Forms.Button();
            this.buttonWorkbook = new System.Windows.Forms.Button();
            this.buttonExcel = new System.Windows.Forms.Button();
            this.labelProxyCount = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonAddRemoveWorkbook
            // 
            this.buttonAddRemoveWorkbook.Enabled = false;
            this.buttonAddRemoveWorkbook.Location = new System.Drawing.Point(369, 25);
            this.buttonAddRemoveWorkbook.Name = "buttonAddRemoveWorkbook";
            this.buttonAddRemoveWorkbook.Size = new System.Drawing.Size(176, 25);
            this.buttonAddRemoveWorkbook.TabIndex = 20;
            this.buttonAddRemoveWorkbook.Text = "Add && Remove Workbook";
            this.buttonAddRemoveWorkbook.UseVisualStyleBackColor = true;
            this.buttonAddRemoveWorkbook.Click += new System.EventHandler(this.buttonAddRemoveWorkbook_Click);
            // 
            // buttonAddins
            // 
            this.buttonAddins.Enabled = false;
            this.buttonAddins.Location = new System.Drawing.Point(260, 25);
            this.buttonAddins.Name = "buttonAddins";
            this.buttonAddins.Size = new System.Drawing.Size(103, 25);
            this.buttonAddins.TabIndex = 19;
            this.buttonAddins.Text = "Enum Addins";
            this.buttonAddins.UseVisualStyleBackColor = true;
            this.buttonAddins.Click += new System.EventHandler(this.buttonAddins_Click);
            // 
            // buttonWorkbook
            // 
            this.buttonWorkbook.Enabled = false;
            this.buttonWorkbook.Location = new System.Drawing.Point(151, 25);
            this.buttonWorkbook.Name = "buttonWorkbook";
            this.buttonWorkbook.Size = new System.Drawing.Size(103, 25);
            this.buttonWorkbook.TabIndex = 18;
            this.buttonWorkbook.Text = "Add Workbook";
            this.buttonWorkbook.UseVisualStyleBackColor = true;
            this.buttonWorkbook.Click += new System.EventHandler(this.buttonWorkbook_Click);
            // 
            // buttonExcel
            // 
            this.buttonExcel.Location = new System.Drawing.Point(32, 25);
            this.buttonExcel.Name = "buttonExcel";
            this.buttonExcel.Size = new System.Drawing.Size(103, 25);
            this.buttonExcel.TabIndex = 17;
            this.buttonExcel.Text = "Start Excel";
            this.buttonExcel.UseVisualStyleBackColor = true;
            this.buttonExcel.Click += new System.EventHandler(this.buttonExcel_Click);
            // 
            // labelProxyCount
            // 
            this.labelProxyCount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelProxyCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProxyCount.Location = new System.Drawing.Point(237, 90);
            this.labelProxyCount.Name = "labelProxyCount";
            this.labelProxyCount.Size = new System.Drawing.Size(47, 20);
            this.labelProxyCount.TabIndex = 16;
            this.labelProxyCount.Text = "0";
            this.labelProxyCount.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(34, 90);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(197, 20);
            this.label1.TabIndex = 15;
            this.label1.Text = "Current COM Proxies open";
            // 
            // Tutorial04
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.buttonAddRemoveWorkbook);
            this.Controls.Add(this.buttonAddins);
            this.Controls.Add(this.buttonWorkbook);
            this.Controls.Add(this.buttonExcel);
            this.Controls.Add(this.labelProxyCount);
            this.Controls.Add(this.label1);
            this.Name = "Tutorial04";
            this.Size = new System.Drawing.Size(686, 478);
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
    }
}
