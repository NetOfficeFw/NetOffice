using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.Controls.Error;

namespace NetOffice.DeveloperToolbox.Forms
{
    partial class ErrorForm : Form
    {
        #region Construction

        public ErrorForm()
        {
            InitializeComponent();
        }

        public ErrorForm(Exception exception, string message, ErrorCategory category, int currentLanguageID)
        {
            InitializeComponent();
            errorControl1.ShowError(exception, message, category, currentLanguageID);
        }

        public ErrorForm(Exception exception, ErrorCategory category, int currentLanguageID)
        {
            InitializeComponent();
            errorControl1.ShowError(exception, category, currentLanguageID);
        }

        #endregion

        #region Methods

        public static void ShowError(Exception exception, ErrorCategory category, int currentLanguageID)
        {
            ErrorForm form = new ErrorForm(exception, category, currentLanguageID);
            if (null != MainForm.Singleton && MainForm.Singleton.Visible)
                form.ShowDialog(MainForm.Singleton);
            else
            {
                form.StartPosition = FormStartPosition.CenterScreen;
                form.ShowDialog();
            }
        }

        public static void ShowError(IWin32Window parent, Exception exception, ErrorCategory category, int currentLanguageID)
        {
            ErrorForm form = new ErrorForm(exception, category, currentLanguageID);
            form.ShowDialog(parent);
        }

        public static void ShowError(IWin32Window parent, Exception exception, ErrorCategory category)
        {
            ErrorForm form = new ErrorForm(exception, category, Forms.MainForm.Singleton.CurrentLanguageID );
            form.ShowDialog(parent);
        }

        public static void ShowError(IWin32Window parent, Exception exception)
        {
            ErrorForm form = new ErrorForm(exception, ErrorCategory.NonCritical, Forms.MainForm.Singleton.CurrentLanguageID);
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
