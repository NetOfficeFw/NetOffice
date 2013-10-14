using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    public class StaticOpenFilePanelDialogResult
    {
        internal StaticOpenFilePanelDialogResult(DialogResult result, string[] files)
        {
            Result = result;
            SelectedFiles = files;
        }
        public DialogResult Result { get; private set; }
        public string[] SelectedFiles { get; private set; }
    }
}
