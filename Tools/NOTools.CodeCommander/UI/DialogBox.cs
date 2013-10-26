using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace NOTools.CodeCommander.UI
{
    /// <summary>
    /// helper: display messages and errors
    /// </summary>
    internal static class DialogBox
    {
        /// <summary>
        /// shows a simple message in a messagebox
        /// </summary>
        /// <param name="message">message to display</param>
        public static void ShowMessage(string message)
        {
            MessageBox.Show(message, "NetOfficeDeveloperAddin", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// shows a simple error message in a messagebox
        /// </summary>
        /// <param name="message">errormessage to display</param>
        public static void ShowSimpleError(string message)
        {
            MessageBox.Show(message, "NetOfficeDeveloperAddin", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// shows an error with detailed exception message
        /// </summary>
        /// <param name="methodName">name of the error method</param>
        /// <param name="exception">throwed exception</param>
        public static void ShowDetailedError(string methodName, Exception exception)
        {
            string message = string.Format("An error ocurred in {2}{1}{1}{1}Details:{1}{1}{0}", exception.Message, Environment.NewLine, methodName);
            MessageBox.Show(message, "NetOfficeDeveloperAddin", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
