using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace TutorialsBase
{
    /// <summary>
    /// the host application for the examples
    /// </summary>
    public interface IHost
    {
        /// <summary>
        /// shows finished tutorial dialog
        /// </summary>
        void ShowFinishDialog();

        /// <summary>
        /// shows finished tutorial dialog with given message.
        /// </summary>
        /// <param name="message">given message as any</param>
        void ShowFinishDialog(string message);

        /// <summary>
        /// shows an error dialog
        /// </summary>
        /// <param name="message">given message as any</param>
        /// <param name="exception">given exception as any</param>
        void ShowErrorDialog(string message, Exception exception);

        /// <summary>
        /// shows a question
        /// </summary>
        /// <param name="message">question to the user</param>
        /// <returns>yes or no</returns>
        DialogResult ShowQuestion(string message);

        /// <summary>
        /// shows a message
        /// </summary>
        /// <param name="message">message as information</param>
        void ShowMessage(string message);

        /// <summary>
        /// the host application select a tutorial
        /// </summary>
        /// <param name="index">zero based index from the target tutorial</param>
        void NavigateToTutorial(int index);

        /// <summary>
        /// Helper icon for some examples
        /// </summary>
        Icon DisplayIcon { get; }

        /// <summary>
        /// Current Language. only english or german (1033 or 1031)
        /// </summary>
        int LCID { get; }
        
    }
}
