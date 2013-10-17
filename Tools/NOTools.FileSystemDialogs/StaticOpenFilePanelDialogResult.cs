using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// Result from OpenFilePanelDialog static Show methd
    /// </summary>
    public class StaticOpenFilePanelDialogResult
    {
        internal StaticOpenFilePanelDialogResult(DialogResult result, string[] files)
        {
            Result = result;
            SelectedFiles = files;
        }

        /// <summary>
        /// User want complete the operation(Ok) or cancel(Cancel)
        /// </summary>
        public DialogResult Result { get; private set; }

        /// <summary>
        /// User selected files
        /// </summary>
        public string[] SelectedFiles { get; private set; }
    }
}
