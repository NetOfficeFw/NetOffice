using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    public class FileDoubleClickEventArgs : EventArgs
    {
        internal FileDoubleClickEventArgs(string file)
        {
            File = file;
        }
        public string File{ get; private set; }
    }

    public delegate void FileDoubleClickEventHandler(object sender, FileDoubleClickEventArgs args);
}
