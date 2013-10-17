using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// FileDoubleClick event arguments
    /// </summary>
    public class FileDoubleClickEventArgs : EventArgs
    {
        internal FileDoubleClickEventArgs(string file)
        {
            File = file;
        }

        /// <summary>
        /// User Selected File
        /// </summary>
        public string File{ get; private set; }
    }

    public delegate void FileDoubleClickEventHandler(object sender, FileDoubleClickEventArgs args);
}
