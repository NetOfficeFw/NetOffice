namespace NetOffice.DeveloperToolbox.Translation
{
    partial class ToolLanguageForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ToolLanguageForm));
            this.buttonSaveChanges = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.label1DefaultHint = new System.Windows.Forms.Label();
            this.pictureBox6 = new System.Windows.Forms.PictureBox();
            this.panelDefaultHint = new System.Windows.Forms.Panel();
            this.toolLanguageControl1 = new NetOffice.DeveloperToolbox.Translation.ToolLanguageControl();
            this.overlayPainter1 = new NetOffice.DeveloperToolbox.Controls.Painter.OverlayPainter(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox6)).BeginInit();
            this.panelDefaultHint.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonSaveChanges
            // 
            resources.ApplyResources(this.buttonSaveChanges, "buttonSaveChanges");
            this.buttonSaveChanges.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonSaveChanges.ForeColor = System.Drawing.Color.Blue;
            this.buttonSaveChanges.Name = "buttonSaveChanges";
            this.buttonSaveChanges.UseVisualStyleBackColor = true;
            this.buttonSaveChanges.Click += new System.EventHandler(this.buttonSaveChanges_Click);
            // 
            // buttonCancel
            // 
            resources.ApplyResources(this.buttonCancel, "buttonCancel");
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonCancel.ForeColor = System.Drawing.Color.Blue;
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // label1DefaultHint
            // 
            resources.ApplyResources(this.label1DefaultHint, "label1DefaultHint");
            this.label1DefaultHint.ForeColor = System.Drawing.Color.White;
            this.label1DefaultHint.Name = "label1DefaultHint";
            // 
            // pictureBox6
            // 
            resources.ApplyResources(this.pictureBox6, "pictureBox6");
            this.pictureBox6.Name = "pictureBox6";
            this.pictureBox6.TabStop = false;
            // 
            // panelDefaultHint
            // 
            resources.ApplyResources(this.panelDefaultHint, "panelDefaultHint");
            this.panelDefaultHint.Controls.Add(this.label1DefaultHint);
            this.panelDefaultHint.Controls.Add(this.pictureBox6);
            this.panelDefaultHint.Name = "panelDefaultHint";
            // 
            // toolLanguageControl1
            // 
            resources.ApplyResources(this.toolLanguageControl1, "toolLanguageControl1");
            this.toolLanguageControl1.Name = "toolLanguageControl1";
            this.toolLanguageControl1.SelectedNodeTextChanged += new System.EventHandler(this.toolLanguageControl1_SelectedNodeTextChanged);
            this.toolLanguageControl1.SelectedTabChanged += new System.EventHandler(this.toolLanguageControl1_SelectedTabChanged);
            // 
            // overlayPainter1
            // 
            this.overlayPainter1.Paint += new System.EventHandler<System.Windows.Forms.PaintEventArgs>(this.overlayPainter1_Paint);
            // 
            // ToolLanguageForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.CancelButton = this.buttonCancel;
            this.Controls.Add(this.panelDefaultHint);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonSaveChanges);
            this.Controls.Add(this.toolLanguageControl1);
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ToolLanguageForm";
            this.ShowInTaskbar = false;
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ToolLanguageForm_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox6)).EndInit();
            this.panelDefaultHint.ResumeLayout(false);
            this.panelDefaultHint.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private ToolLanguageControl toolLanguageControl1;
        private System.Windows.Forms.Button buttonSaveChanges;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Label label1DefaultHint;
        private System.Windows.Forms.PictureBox pictureBox6;
        private System.Windows.Forms.Panel panelDefaultHint;
        private Controls.Painter.OverlayPainter overlayPainter1;
    }
}