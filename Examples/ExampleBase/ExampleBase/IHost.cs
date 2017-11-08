using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExampleBase
{
    /// <summary>
    /// Represents the host application
    /// </summary>
    public interface IHost
    {
        /// <summary>
        /// Shows a dialog with the given message. provides an "Open Document" button for the second argument.
        /// </summary>
        /// <param name="message">message as any</param>
        /// <param name="fullDocumentPath">path to result document</param>
        void ShowFinishDialog(string message, string fullDocumentPath);

        /// <summary>
        /// Shows an error dialog
        /// </summary>
        /// <param name="message">message as any</param>
        /// <param name="exception">exception as any</param>
        void ShowErrorDialog(string message, Exception exception);

        /// <summary>
        /// Shows a message
        /// </summary>
        /// <param name="message">message as any</param>
        void ShowMessage(string message);

        /// <summary>
        /// Shows a question
        /// </summary>
        /// <param name="message">message as any</param>
        /// <returns>user response</returns>
        DialogResult ShowQuestion(string message);

        /// <summary>
        /// A dumy icon that examples can use
        /// </summary>
        Icon DisplayIcon { get; }

        /// <summary>
        /// Basefolder path to generate examples documents into
        /// </summary>
        string RootDirectory { get; }
    }
}