using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    public class SelectionChangedEventArgs : EventArgs
    {
        internal SelectionChangedEventArgs(string[] files)
        {
            Files = files;
        }
        public string[] Files { get; private set; }
    }

    public delegate void SelectionChangedEventHandler(object sender, SelectionChangedEventArgs args);
}
