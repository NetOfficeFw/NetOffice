namespace NOTools.DeveloperAddin.UI
{
    partial class PropertyPane
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
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.comboBoxTarget = new System.Windows.Forms.ComboBox();
            this.propertyGridHostProperties = new System.Windows.Forms.PropertyGrid();
            this.labelTargetCaption = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRefresh.Location = new System.Drawing.Point(4, 447);
            this.buttonRefresh.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new System.Drawing.Size(408, 30);
            this.buttonRefresh.TabIndex = 10;
            this.buttonRefresh.Text = "Refresh";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // comboBoxTarget
            // 
            this.comboBoxTarget.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxTarget.BackColor = System.Drawing.Color.LightSteelBlue;
            this.comboBoxTarget.DisplayMember = "Name";
            this.comboBoxTarget.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxTarget.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBoxTarget.Location = new System.Drawing.Point(85, 7);
            this.comboBoxTarget.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.comboBoxTarget.Name = "comboBoxTarget";
            this.comboBoxTarget.Size = new System.Drawing.Size(324, 24);
            this.comboBoxTarget.TabIndex = 8;
            this.comboBoxTarget.ValueMember = "ID";
            this.comboBoxTarget.SelectedValueChanged += new System.EventHandler(this.comboBoxTarget_SelectedValueChanged);
            // 
            // propertyGridHostProperties
            // 
            this.propertyGridHostProperties.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.propertyGridHostProperties.CommandsBackColor = System.Drawing.Color.White;
            this.propertyGridHostProperties.CommandsDisabledLinkColor = System.Drawing.Color.RoyalBlue;
            this.propertyGridHostProperties.HelpVisible = false;
            this.propertyGridHostProperties.LineColor = System.Drawing.Color.White;
            this.propertyGridHostProperties.Location = new System.Drawing.Point(4, 41);
            this.propertyGridHostProperties.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.propertyGridHostProperties.Name = "propertyGridHostProperties";
            this.propertyGridHostProperties.PropertySort = System.Windows.Forms.PropertySort.Alphabetical;
            this.propertyGridHostProperties.Size = new System.Drawing.Size(408, 404);
            this.propertyGridHostProperties.TabIndex = 7;
            this.propertyGridHostProperties.ToolbarVisible = false;
            this.propertyGridHostProperties.ViewBackColor = System.Drawing.Color.LightSteelBlue;
            // 
            // labelTargetCaption
            // 
            this.labelTargetCaption.AutoSize = true;
            this.labelTargetCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTargetCaption.Location = new System.Drawing.Point(1, 11);
            this.labelTargetCaption.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelTargetCaption.Name = "labelTargetCaption";
            this.labelTargetCaption.Size = new System.Drawing.Size(58, 16);
            this.labelTargetCaption.TabIndex = 9;
            this.labelTargetCaption.Text = "Choose:";
            // 
            // PropertyPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.comboBoxTarget);
            this.Controls.Add(this.propertyGridHostProperties);
            this.Controls.Add(this.labelTargetCaption);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "PropertyPane";
            this.Size = new System.Drawing.Size(416, 478);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonRefresh;
        private System.Windows.Forms.ComboBox comboBoxTarget;
        private System.Windows.Forms.PropertyGrid propertyGridHostProperties;
        private System.Windows.Forms.Label labelTargetCaption;
    }
}
