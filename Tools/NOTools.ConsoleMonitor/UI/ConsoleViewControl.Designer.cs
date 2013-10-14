namespace NOTools.ConsoleMonitor
{
    partial class ConsoleViewControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConsoleViewControl));
            this.radioButtonHierarchy = new System.Windows.Forms.RadioButton();
            this.radioButtonPlain = new System.Windows.Forms.RadioButton();
            this.labelViewStyle = new System.Windows.Forms.Label();
            this.TextBoxConsole = new System.Windows.Forms.TextBox();
            this.buttonCloseConsole = new System.Windows.Forms.Button();
            this.radioButtonPlainReverse = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // radioButtonHierarchy
            // 
            this.radioButtonHierarchy.AutoSize = true;
            this.radioButtonHierarchy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonHierarchy.Location = new System.Drawing.Point(257, 12);
            this.radioButtonHierarchy.Name = "radioButtonHierarchy";
            this.radioButtonHierarchy.Size = new System.Drawing.Size(84, 20);
            this.radioButtonHierarchy.TabIndex = 10;
            this.radioButtonHierarchy.Text = "Hierarchy";
            this.radioButtonHierarchy.UseVisualStyleBackColor = true;
            this.radioButtonHierarchy.CheckedChanged += new System.EventHandler(this.radioButtonViewStyle_CheckedChanged);
            // 
            // radioButtonPlain
            // 
            this.radioButtonPlain.AutoSize = true;
            this.radioButtonPlain.Checked = true;
            this.radioButtonPlain.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonPlain.Location = new System.Drawing.Point(82, 10);
            this.radioButtonPlain.Name = "radioButtonPlain";
            this.radioButtonPlain.Size = new System.Drawing.Size(56, 20);
            this.radioButtonPlain.TabIndex = 9;
            this.radioButtonPlain.TabStop = true;
            this.radioButtonPlain.Text = "Plain";
            this.radioButtonPlain.UseVisualStyleBackColor = true;
            this.radioButtonPlain.CheckedChanged += new System.EventHandler(this.radioButtonViewStyle_CheckedChanged);
            // 
            // labelViewStyle
            // 
            this.labelViewStyle.AutoSize = true;
            this.labelViewStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelViewStyle.Location = new System.Drawing.Point(8, 12);
            this.labelViewStyle.Name = "labelViewStyle";
            this.labelViewStyle.Size = new System.Drawing.Size(70, 16);
            this.labelViewStyle.TabIndex = 8;
            this.labelViewStyle.Text = "ViewStyle:";
            // 
            // TextBoxConsole
            // 
            this.TextBoxConsole.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxConsole.BackColor = System.Drawing.SystemColors.Window;
            this.TextBoxConsole.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TextBoxConsole.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TextBoxConsole.Location = new System.Drawing.Point(0, 40);
            this.TextBoxConsole.Multiline = true;
            this.TextBoxConsole.Name = "TextBoxConsole";
            this.TextBoxConsole.ReadOnly = true;
            this.TextBoxConsole.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.TextBoxConsole.Size = new System.Drawing.Size(640, 440);
            this.TextBoxConsole.TabIndex = 13;
            this.TextBoxConsole.WordWrap = false;
            // 
            // buttonCloseConsole
            // 
            this.buttonCloseConsole.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCloseConsole.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonCloseConsole.Image = ((System.Drawing.Image)(resources.GetObject("buttonCloseConsole.Image")));
            this.buttonCloseConsole.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonCloseConsole.Location = new System.Drawing.Point(483, 6);
            this.buttonCloseConsole.Name = "buttonCloseConsole";
            this.buttonCloseConsole.Size = new System.Drawing.Size(140, 29);
            this.buttonCloseConsole.TabIndex = 14;
            this.buttonCloseConsole.Text = " Close Console";
            this.buttonCloseConsole.UseVisualStyleBackColor = true;
            this.buttonCloseConsole.Visible = false;
            this.buttonCloseConsole.Click += new System.EventHandler(this.buttonCloseConsole_Click);
            // 
            // radioButtonPlainReverse
            // 
            this.radioButtonPlainReverse.AutoSize = true;
            this.radioButtonPlainReverse.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonPlainReverse.Location = new System.Drawing.Point(140, 11);
            this.radioButtonPlainReverse.Name = "radioButtonPlainReverse";
            this.radioButtonPlainReverse.Size = new System.Drawing.Size(111, 20);
            this.radioButtonPlainReverse.TabIndex = 15;
            this.radioButtonPlainReverse.Text = "Plain Reverse";
            this.radioButtonPlainReverse.UseVisualStyleBackColor = true;
            this.radioButtonPlainReverse.CheckedChanged += new System.EventHandler(this.radioButtonViewStyle_CheckedChanged);
            // 
            // ConsoleViewControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.radioButtonPlainReverse);
            this.Controls.Add(this.buttonCloseConsole);
            this.Controls.Add(this.TextBoxConsole);
            this.Controls.Add(this.radioButtonHierarchy);
            this.Controls.Add(this.radioButtonPlain);
            this.Controls.Add(this.labelViewStyle);
            this.Name = "ConsoleViewControl";
            this.Size = new System.Drawing.Size(640, 480);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton radioButtonHierarchy;
        private System.Windows.Forms.RadioButton radioButtonPlain;
        private System.Windows.Forms.Label labelViewStyle;
        private System.Windows.Forms.TextBox TextBoxConsole;
        private System.Windows.Forms.Button buttonCloseConsole;
        private System.Windows.Forms.RadioButton radioButtonPlainReverse;
    }
}
