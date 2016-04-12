using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using NetOffice.OfficeApi.Tools.Informations;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Shows essential technical environment information to the user 
    /// </summary>
    public partial class DiagnosticsDialog : ToolsDialog
    {
        #region Fields

        #endregion

        private const string _assemblyInfoTemplate = "%AssemblyInfo";

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public DiagnosticsDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        ///  Creates an instance of the class
        /// </summary>
        /// <param name="diagnostics">diagnostics to display</param>
        public DiagnosticsDialog(IEnumerable<DiagnosticPair> diagnostics)
        {
            InitializeComponent();
            dataGridViewDiagnostics.AutoGenerateColumns = false;
            dataGridViewDiagnostics.DataSource = diagnostics;
        }

        #endregion

        #region Methods

        private string CreateClipboardContent()
        {
            StringBuilder sb = new StringBuilder();
            IEnumerable<DiagnosticPair> dataSource = dataGridViewDiagnostics.DataSource as IEnumerable<DiagnosticPair>;
            if (null != dataSource)
            {
                foreach (DiagnosticPair item in dataSource)
                    sb.AppendLine(String.Format("Type:{0} Value:{1}", item.Type, item.Value));
            }
            else
            {
                sb.AppendLine("NetOffice Diagnostics:<Empty>");
            }

            return sb.ToString();
        }

        #endregion

        #region Overrides

        /// <summary>
        /// <see cref="ToolsDialog.DoLocalization"/>
        /// </summary>
        /// <param name="localization">localized values</param>
        protected internal override void DoLocalization(DialogLocalization localization)
        {
            Text = localization["Title", Text];
            labelAssemblyInfo.Text = localization["labelAssemblyInfo", labelAssemblyInfo.Text];
            buttonClipboardCopy.Text = localization["buttonClipboardCopy", buttonClipboardCopy.Text];
            buttonClose.Text = localization["buttonClose", buttonClose.Text];
            colType.HeaderText = localization["Type", colType.HeaderText];
            colValue.HeaderText = localization["Value", colValue.HeaderText];
        }

        /// <summary>
        /// <see cref="ToolsDialog.DoLayout"/>
        /// </summary>
        /// <param name="layout">layout settings</param>
        protected internal override void DoLayout(DialogLayoutSettings layout)
        {
            dataGridViewDiagnostics.BackgroundColor = layout.BackHeaderColor;
            dataGridViewDiagnostics.ColumnHeadersDefaultCellStyle.BackColor = layout.BackColor;
            dataGridViewDiagnostics.ColumnHeadersDefaultCellStyle.ForeColor = layout.ForeAlternateColor;
            base.DoLayout(layout);
        }

        #endregion

        #region Trigger

        private void buttonClipboardCopy_Click(object sender, EventArgs e)
        {
            try
            {
                string content = CreateClipboardContent();
                Clipboard.SetText(content);
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch(Exception exception)
            {
                ShowSingleException(exception);
            }
        }

        #endregion
    }
}
