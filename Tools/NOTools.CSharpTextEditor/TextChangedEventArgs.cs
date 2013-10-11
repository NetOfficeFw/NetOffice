using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    public class TextChangedEventArgs : EventArgs
    {
        public string Text { get; private set; }

        internal TextChangedEventArgs(string text)
        {
            Text = text;
        }
    }

    public delegate void TextChangedEventHander(object sender, TextChangedEventArgs args);

}
