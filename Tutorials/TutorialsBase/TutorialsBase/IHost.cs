using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace TutorialsBase
{
    /// <summary>
    /// the host application for the examples
    /// </summary>
    public interface IHost
    {
        /// <summary>
        /// shows a dialog with the given message.
        /// </summary>
        /// <param name="message"></param>
        void ShowFinishDialog();
        void ShowFinishDialog(string message);

        /// <summary>
        /// shows an error dialog
        /// </summary>
        /// <param name="message"></param>
        /// <param name="exception"></param>
        void ShowErrorDialog(string message, Exception exception);

        /// <summary>
        /// the host application select a tutorial
        /// </summary>
        /// <param name="index"></param>
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
