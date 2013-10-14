using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    public class CompileRequestEventArgs : EventArgs
    {
        internal CompileRequestEventArgs(Key key)
        {
            Key = key;
        }

        public Key Key { get; private set; }
    }

    public delegate void CompileRequestHandler(CodeEditorControl sender, CompileRequestEventArgs args);
}
