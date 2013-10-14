using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.CodeDom.Compiler;

namespace NOTools.CSharpTextEditor
{
    /// <summary>
    /// Show compiler error messages in a datagrid
    /// </summary>
    internal partial class ErrorPanel : UserControl
    {
        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ErrorPanel()
        {
            InitializeComponent();
            dataGridErrors.AutoGenerateColumns = false;
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs on double click for a row
        /// </summary>
        public event ErrorDoubleClickEventHandler ErrorDoubleClick;

        private void RaiseErrorDoubleClick(int lineNumber, int columnNumber)
        {
            if (null != ErrorDoubleClick)
                ErrorDoubleClick(this, lineNumber, columnNumber);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Column header for the first column
        /// </summary>
        public string LineColumnHeader
        {
            get
            {
                return ColumnLineNumber.HeaderText;
            }
            set
            {
                ColumnLineNumber.HeaderText = value;
            }
        }

        /// <summary>
        /// Column header for the second column
        /// </summary>
        public string ErrorColumnHeader
        {
            get
            {
                return ColumnMessage.HeaderText;
            }
            set
            {
                ColumnMessage.HeaderText = value;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Show all compiler erros in a datagrid
        /// </summary>
        public void ShowErrors(CompilerErrorCollection errors)
        {
            dataGridErrors.DataSource = errors;
        }

        /// <summary>
        ///  Clear the data grid
        /// </summary>
        public void ClearErrors()
        {
            dataGridErrors.DataSource = null;
        }

        /// <summary>
        /// Set the column header backcolor
        /// </summary>
        /// <param name="color">new color</param>
        public void SetColumnColor(Color color)
        {
            dataGridErrors.ColumnHeadersDefaultCellStyle.BackColor = color;
        }

        #endregion

        #region Trigger

        private void dataGridErrors_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow currentRow = dataGridErrors.Rows[e.RowIndex];
            if (null != currentRow.Cells[0].Value && null != currentRow.Cells[1].Value)
            { 
                int lineNumber =  (int)currentRow.Cells[0].Value;
                int columnNumber = (int)currentRow.Cells[1].Value;
                RaiseErrorDoubleClick(lineNumber, columnNumber);
            }
        }

        #endregion
    }
}
