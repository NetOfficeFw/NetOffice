using System;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Forms
{
    /// <summary>
    /// Application Error Form
    /// </summary>
    partial class ErrorForm : Form
    {
        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ErrorForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="exception">exception as any</param>
        /// <param name="category">error category</param>
        public ErrorForm(Exception exception, ErrorCategory category)
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
            errorControl1.ShowError(exception, category);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="exception">exception as any</param>
        /// <param name="message">message to show</param>
        /// <param name="category">error category</param>
        public ErrorForm(Exception exception,  string message, ErrorCategory category)
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
            errorControl1.ShowError(exception, message, category);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates an instance of ErrorForm and show
        /// </summary>
        /// <param name="parent">modal parent</param>
        /// <param name="exception">exception as any</param>
        /// <param name="category">error category</param>
        /// <param name="currentLanguageID">current user language</param>
        public static void ShowError(IWin32Window parent, Exception exception, ErrorCategory category)
        {
            ErrorForm form = new ErrorForm(exception, category);
            form.ShowDialog(parent);
        }

        /// <summary>
        /// Creates an instance of ErrorForm and show
        /// </summary>
        /// <param name="parent">modal parent</param>
        /// <param name="exception">exception as any</param>
        /// <param name="category">error category</param>
        /// <param name="message">friendly header message for the user</param>
        public static void ShowError(IWin32Window parent, Exception exception, ErrorCategory category, string message)
        {
            ErrorForm form = new ErrorForm(exception, message, category);
            form.ShowDialog(parent);
        }

        /// <summary>
        /// Creates an instance of ErrorForm and show
        /// </summary>
        /// <param name="parent">modal parent</param>
        /// <param name="exception">exception as any</param>
        public static void ShowError(IWin32Window parent, Exception exception)
        {
            ErrorForm form = new ErrorForm(exception, ErrorCategory.NonCritical);
            form.ShowDialog(parent);
        }

        #endregion

        #region Trigger

        private void errorControl1_UserClose(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion
    }
}
