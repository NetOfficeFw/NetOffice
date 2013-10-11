namespace NOTools.CSharpTextEditor
{
    partial class ErrorPanel
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridErrors = new System.Windows.Forms.DataGridView();
            this.ColumnLineNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnMessage = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridErrors)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridErrors
            // 
            this.dataGridErrors.AllowUserToAddRows = false;
            this.dataGridErrors.AllowUserToDeleteRows = false;
            this.dataGridErrors.AllowUserToResizeColumns = false;
            this.dataGridErrors.BackgroundColor = System.Drawing.Color.White;
            this.dataGridErrors.BorderStyle = System.Windows.Forms.BorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.LightSteelBlue;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridErrors.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridErrors.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridErrors.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnLineNumber,
            this.ColumnColumn,
            this.ColumnMessage});
            this.dataGridErrors.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridErrors.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridErrors.GridColor = System.Drawing.Color.Black;
            this.dataGridErrors.Location = new System.Drawing.Point(0, 0);
            this.dataGridErrors.MultiSelect = false;
            this.dataGridErrors.Name = "dataGridErrors";
            this.dataGridErrors.RowHeadersVisible = false;
            this.dataGridErrors.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridErrors.ShowCellErrors = false;
            this.dataGridErrors.ShowEditingIcon = false;
            this.dataGridErrors.ShowRowErrors = false;
            this.dataGridErrors.Size = new System.Drawing.Size(627, 190);
            this.dataGridErrors.TabIndex = 0;
            this.dataGridErrors.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridErrors_CellDoubleClick);
            // 
            // ColumnLineNumber
            // 
            this.ColumnLineNumber.DataPropertyName = "Line";
            this.ColumnLineNumber.HeaderText = "Line";
            this.ColumnLineNumber.Name = "ColumnLineNumber";
            this.ColumnLineNumber.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.ColumnLineNumber.Width = 70;
            // 
            // ColumnColumn
            // 
            this.ColumnColumn.DataPropertyName = "Column";
            this.ColumnColumn.HeaderText = "Column";
            this.ColumnColumn.Name = "ColumnColumn";
            this.ColumnColumn.Visible = false;
            // 
            // ColumnMessage
            // 
            this.ColumnMessage.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ColumnMessage.DataPropertyName = "ErrorText";
            this.ColumnMessage.HeaderText = "Error";
            this.ColumnMessage.Name = "ColumnMessage";
            // 
            // ErrorPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.dataGridErrors);
            this.Name = "ErrorPanel";
            this.Size = new System.Drawing.Size(627, 190);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridErrors)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridErrors;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnLineNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnMessage;
    }
}
