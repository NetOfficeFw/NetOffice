using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace ExampleBase
{
    /// <summary>
    /// the host application for the examples
    /// </summary>
    public interface IHost
    {
        /// <summary>
        /// shows a dialog with the given message. provides an "Open Document" button for the second param. this param means a full qualified path to a generated file.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="fullDocumentPath"></param>
        void ShowFinishDialog(string message, string fullDocumentPath);

        /// <summary>
        /// shows an error dialog
        /// </summary>
        /// <param name="message"></param>
        /// <param name="exception"></param>
        void ShowErrorDialog(string message, Exception exception);

        /// <summary>
        /// Helper icon for some examples
        /// </summary>
        Icon DisplayIcon { get; }

        /// <summary>
        /// Current Language. only english or german (1033 or 1031)
        /// </summary>
        int LCID { get; }

        /// <summary>
        /// Basefolder information for  document generating examples 
        /// </summary>
        string RootDirectory { get; }
    }
}
