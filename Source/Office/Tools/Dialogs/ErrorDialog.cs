using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Shows error information to the user
    /// </summary>
    public partial class ErrorDialog : ToolsDialog
    {
        #region Embeded Definitions

        /// <summary>
        /// Encapsulate an exception to show them as easy
        /// </summary>
        private class ErrorDescription
        {
            #region Fields

            private Exception _exception;

            #endregion

            #region Ctor

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="exception">top level exception to display</param>
            internal ErrorDescription(Exception exception)
            {
                _exception = exception;
                Message =  null != exception ? exception.Message : String.Empty;
                Type = null != exception ? exception.GetType().Name : String.Empty;
                Source = null != exception && null != exception.TargetSite ? exception.TargetSite.ToString() : String.Empty;
            }

            #endregion

            #region Properties

            /// <summary>
            /// Exception Message
            /// </summary>
            public string Message { get; private set; }

            /// <summary>
            /// Type of Exception
            /// </summary>
            public string Type { get; private set; }

            /// <summary>
            /// Source/Scope from the Exception
            /// </summary>
            public string Source { get; private set; }

            #endregion

            #region Methods

            /// <summary>
            /// Create enumerator for given exception and all inner exception
            /// </summary>
            /// <param name="exception">last exception in stack</param>
            /// <returns>enumerator to iterate the errors</returns>
            internal static IEnumerable<ErrorDescription> CreateList(Exception exception)
            {
                List<ErrorDescription> list = new List<ErrorDescription>();

                while (null != exception)
                {
                    ErrorDescription error = new ErrorDescription(exception);
                    list.Add(error);
                    exception = exception.InnerException;
                }

                return list;
            }

            #endregion

            #region Overrides

            /// <summary>
            /// Returns a System.String that represents the instance
            /// </summary>
            /// <returns>System.String</returns>
            public override string ToString()
            {
                return null != _exception ? _exception.ToString() : base.ToString();
            }

            #endregion
        }

        #endregion

        #region Fields

        private const int _minimumWidth = 540;
        private const int _smallHeight = 210;
        private const int _extendedHeight = 500;
        private const string _errorMessageTemplate = "%ErrorMessage";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ErrorDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="exception">thrown exception to display</param>
        /// <param name="errorMessage">friendly error message to explain what happen</param>
        /// <param name="allowDetails">allow user to see exception details</param>
        public ErrorDialog(Exception exception, string errorMessage, bool allowDetails)
        {
            InitializeComponent();
            buttonShowDetails.Visible = allowDetails;
            Height = _smallHeight;
            labelErrorMessage.Text = _errorMessageTemplate;
            dataGridViewErrors.AutoGenerateColumns = false;
            if (!String.IsNullOrEmpty(errorMessage))
                labelErrorMessage.Text = errorMessage;
            dataGridViewErrors.DataSource = ErrorDescription.CreateList(exception);
            Height = _smallHeight;
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
            labelErrorHeader.Text = localization["ErrorHeader", labelErrorHeader.Text];
            if (labelErrorMessage.Text.Equals(_errorMessageTemplate, StringComparison.InvariantCultureIgnoreCase))
                labelErrorMessage.Text = localization["ErrorMessage", ""];
            colMessage.HeaderText = localization["Message", colMessage.HeaderText];
            colType.HeaderText = localization["Type", colType.HeaderText];
            colSource.HeaderText = localization["Source", colSource.HeaderText];
            buttonShowDetails.Text = localization["buttonShowDetails", buttonShowDetails.Text];
            buttonClose.Text = localization["buttonClose", buttonClose.Text];
            buttonClipboardCopy.Text = localization["buttonClipboardCopy", buttonClipboardCopy.Text];
        }

        /// <summary>
        /// <see cref="ToolsDialog.DoLayout"/>
        /// </summary>
        /// <param name="layout">layout settings</param>
        protected internal override void DoLayout(DialogLayoutSettings layout)
        {
            dataGridViewErrors.BackgroundColor = layout.BackHeaderColor;
            dataGridViewErrors.ColumnHeadersDefaultCellStyle.BackColor = layout.BackColor;
            dataGridViewErrors.ColumnHeadersDefaultCellStyle.ForeColor = layout.ForeAlternateColor;
            base.DoLayout(layout);
        }

        #endregion

        #region Methods

        private void ShowSingleException(ErrorDescription error)
        {
            MessageBox.Show(this, error.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);        
        }

        private void CopyErrorInfoToClipboard(IEnumerable<ErrorDescription> errors)
        {
            StringBuilder builder = new StringBuilder();

            foreach (ErrorDescription item in errors)
                builder.AppendLine(String.Format("{0};{1};{2}{4}{3}{4}", item.Type, item.Source, item.Message, item.ToString(), Environment.NewLine));

            Clipboard.SetData(DataFormats.Text, builder.ToString());
        }

        #endregion

        #region Trigger

        private void dataGridViewErrors_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dataGridViewErrors.SelectedCells.Count == 0)
                    return;

                ErrorDescription selectedItem = dataGridViewErrors.Rows[dataGridViewErrors.SelectedCells[0].RowIndex].DataBoundItem as ErrorDescription;
                if (null != selectedItem)
                    ShowSingleException(selectedItem);
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);
            }
        }

        private void buttonShowDetails_Click(object sender, EventArgs e)
        {
            try
            {
                buttonShowDetails.Enabled = false;
                FormBorderStyle = FormBorderStyle.Sizable;             
                Height = _extendedHeight;
                MinimumSize = new System.Drawing.Size(_minimumWidth, _extendedHeight);
                buttonShowDetails.Enabled = false;
                dataGridViewErrors.Visible = true;
                buttonClipboardCopy.Visible = true;
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);
            }
        }

        private void buttonClipboardCopy_Click(object sender, EventArgs e)
        {
            try
            {
                IEnumerable<ErrorDescription> list = dataGridViewErrors.DataSource as IEnumerable<ErrorDescription>;
                if (null == list)
                    return;
                CopyErrorInfoToClipboard(list);
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
                Close();
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);
            }
        }

        #endregion
    }
}
