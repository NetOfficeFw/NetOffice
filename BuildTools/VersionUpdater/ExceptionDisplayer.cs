using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text;

namespace NOBuildTools.VersionUpdater
{
    /// <summary>
    /// Exception display helper
    /// </summary>
    internal static class ExceptionDisplayer
    {
        /// <summary>
        /// Shows exception as string message box to the user
        /// </summary>
        /// <param name="parent">modal parent</param>
        /// <param name="exception">exception as any</param>
        public static void ShowException(IWin32Window parent, Exception exception)
        {
            string message = "An error is occured." + Environment.NewLine;
            string detailsMessage = "Details: " + Environment.NewLine;

            while (null != exception)
            {
                detailsMessage += "Exception: " + exception.GetType().Name + Environment.NewLine;
                detailsMessage += "Exception: " + exception.Message + Environment.NewLine;

                exception = exception.InnerException;
            }

            MessageBox.Show(parent, message + detailsMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
