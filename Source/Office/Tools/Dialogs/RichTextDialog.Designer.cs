namespace NetOffice.OfficeApi.Tools.Dialogs
{
    partial class RichTextDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RichTextDialog));
            this.panelHeader = new System.Windows.Forms.Panel();
            this.labelHeaderCaption = new System.Windows.Forms.Label();
            this.pictureBoxHeader = new System.Windows.Forms.PictureBox();
            this.richTextBoxText = new System.Windows.Forms.RichTextBox();
            this.checkBoxCondition = new System.Windows.Forms.CheckBox();
            this.buttonClose = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelHeader
            // 
            this.panelHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelHeader.BackColor = System.Drawing.Color.White;
            this.panelHeader.Controls.Add(this.labelHeaderCaption);
            this.panelHeader.Controls.Add(this.pictureBoxHeader);
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(631, 58);
            this.panelHeader.TabIndex = 2;
            // 
            // labelHeaderCaption
            // 
            this.labelHeaderCaption.AutoSize = true;
            this.labelHeaderCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelHeaderCaption.ForeColor = System.Drawing.Color.Black;
            this.labelHeaderCaption.Location = new System.Drawing.Point(73, 20);
            this.labelHeaderCaption.Name = "labelHeaderCaption";
            this.labelHeaderCaption.Size = new System.Drawing.Size(126, 16);
            this.labelHeaderCaption.TabIndex = 1;
            this.labelHeaderCaption.Text = "%HeaderCaption";
            // 
            // pictureBoxHeader
            // 
            this.pictureBoxHeader.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxHeader.Image")));
            this.pictureBoxHeader.Location = new System.Drawing.Point(25, 13);
            this.pictureBoxHeader.Name = "pictureBoxHeader";
            this.pictureBoxHeader.Size = new System.Drawing.Size(34, 34);
            this.pictureBoxHeader.TabIndex = 0;
            this.pictureBoxHeader.TabStop = false;
            // 
            // richTextBoxText
            // 
            this.richTextBoxText.BackColor = System.Drawing.Color.White;
            this.richTextBoxText.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBoxText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBoxText.Location = new System.Drawing.Point(0, 0);
            this.richTextBoxText.Name = "richTextBoxText";
            this.richTextBoxText.ReadOnly = true;
            this.richTextBoxText.Size = new System.Drawing.Size(630, 311);
            this.richTextBoxText.TabIndex = 3;
            this.richTextBoxText.Text = "";
            // 
            // checkBoxCondition
            // 
            this.checkBoxCondition.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBoxCondition.AutoSize = true;
            this.checkBoxCondition.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxCondition.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxCondition.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxCondition.Location = new System.Drawing.Point(25, 380);
            this.checkBoxCondition.Name = "checkBoxCondition";
            this.checkBoxCondition.Size = new System.Drawing.Size(118, 20);
            this.checkBoxCondition.TabIndex = 4;
            this.checkBoxCondition.Text = "%ConditionText";
            this.checkBoxCondition.UseVisualStyleBackColor = true;
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonClose.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.ForeColor = System.Drawing.Color.Blue;
            this.buttonClose.Image = ((System.Drawing.Image)(resources.GetObject("buttonClose.Image")));
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.Location = new System.Drawing.Point(433, 405);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(176, 29);
            this.buttonClose.TabIndex = 5;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.richTextBoxText);
            this.panel1.Location = new System.Drawing.Point(0, 59);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(632, 313);
            this.panel1.TabIndex = 6;
            // 
            // RichTextDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 453);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.checkBoxCondition);
            this.Controls.Add(this.panelHeader);
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "RichTextDialog";
            this.Text = "Text";
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHeader)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label labelHeaderCaption;
        private System.Windows.Forms.PictureBox pictureBoxHeader;
        private System.Windows.Forms.RichTextBox richTextBoxText;
        private System.Windows.Forms.CheckBox checkBoxCondition;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Panel panel1;
    }
}