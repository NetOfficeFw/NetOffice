using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace NOBuildTools.ReferenceAnalyzer
{
    internal static class ExceptionDisplayer
    {
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
